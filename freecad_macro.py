FONT_SIZE = 6

import FreeCAD, Part, Draft
import FreeCADGui
import math
import re

import sys

def focus_on_all_objects():
    """
    FreeCAD 문서 내 모든 객체를 화면에 표시.
    """
    try:
        view = FreeCADGui.ActiveDocument.ActiveView  # 활성화된 뷰 가져오기
        view.fitAll()  # 모든 객체를 화면에 맞추기
        print("View adjusted to fit all objects.")
    except Exception as e:
        print(f"Error focusing on all objects: {e}")

def stop_macro():
    """
    현재 실행 중인 매크로를 종료합니다.
    """
    print("매크로를 종료합니다.")
    focus_on_all_objects()
    sys.exit()

def wait_for_user():
    """
    사용자 입력 대기 대신 FreeCAD의 이벤트 루프를 강제로 실행하여 대기.
    """
    print("Bodies are displayed in the FreeCAD GUI. Close the message box to continue...")
    from PySide2.QtWidgets import QMessageBox
    box = QMessageBox()
    box.setWindowTitle("Pause")
    box.setText("Press OK to continue...")
    box.exec_()


def clean_line(line):
    """
    대괄호 안의 쉼표는 분리 기준에서 제외하고, 각 필드를 정리하여 반환.
    """
    fields = re.findall(r'\[[^\]]+\]|[^,\t]+', line)
    cleaned_parts = [field.strip() for field in fields]
    return ", ".join(cleaned_parts)

def create_box(center_x, center_y, x_size, y_size, z_start, depth, rotation=0):
    """
    사각형(RECTANGLE) 바디를 생성하여 반환.
    """
    half_x = x_size / 2
    half_y = y_size / 2
    p1 = FreeCAD.Vector(-half_x, -half_y, 0)
    p2 = FreeCAD.Vector(half_x, -half_y, 0)
    p3 = FreeCAD.Vector(half_x, half_y, 0)
    p4 = FreeCAD.Vector(-half_x, half_y, 0)
    wire = Part.makePolygon([p1, p2, p3, p4, p1])
    face = Part.Face(wire)
    solid = face.extrude(FreeCAD.Vector(0, 0, depth))
    solid.translate(FreeCAD.Vector(center_x, center_y, z_start))
    if rotation != 0:
        rotation_axis = FreeCAD.Vector(0, 0, 1)
        solid.rotate(FreeCAD.Vector(center_x, center_y, z_start), rotation_axis, rotation)
    return solid

def create_cylinder(center_x, center_y, radius, z_start, height):
    """
    원(CIRCLE) 바디를 생성하여 반환.
    """
    cylinder = Part.makeCylinder(radius, height, FreeCAD.Vector(center_x, center_y, z_start))
    return cylinder

def debug_body(body, label):
    """
    바디의 BoundBox 정보를 출력하여 디버깅.
    """
    bbox = body.BoundBox
    center_x = (bbox.XMin + bbox.XMax) / 2
    center_y = (bbox.YMin + bbox.YMax) / 2
    x_size = bbox.XMax - bbox.XMin
    y_size = bbox.YMax - bbox.YMin
    print(f"{label} - Center: ({center_x:.3f}, {center_y:.3f}), Size: ({x_size:.3f}, {y_size:.3f})")
    print(f"{label} - BoundBox: Xmin={bbox.XMin}, Xmax={bbox.XMax}, Ymin={bbox.YMin}, Ymax={bbox.YMax}, Zmin={bbox.ZMin}, Zmax={bbox.ZMax}")

def add_text_to_plane(text, position, height=FONT_SIZE, z_position=40, color=(1.0, 0.0, 0.0)):
    """
    Adds text to a specified position in the FreeCAD document with specified color and font size.
    """
    # 텍스트 객체 생성
    text_shape = Draft.make_text(text, FreeCAD.Vector(position[0], position[1], z_position))
    text_shape.Label = f"Text_{text}"  # 객체 라벨 설정
    text_shape.ViewObject.FontSize = height  # 폰트 크기 설정
    text_shape.ViewObject.FontName = "Arial"  # 기본 폰트 설정
    text_shape.ViewObject.TextColor = color  # 텍스트 색상 설정
    return text_shape

def parse_color(color_str):
    """
    색상 문자열을 파싱하여 (R, G, B) 튜플을 반환. (0~1 범위)
    """
    if "ThemeColor" in color_str:
        # ThemeColor를 특정 색상으로 매핑. 필요에 따라 변경 가능.
        return (0.0, 1.0, 0.0)  # 녹색
    numbers = re.findall(r'\d+', color_str)
    if len(numbers) >= 3:
        return tuple(map(lambda x: int(x) / 255.0, numbers[:3]))
    else:
        return (1.0, 1.0, 1.0)  # 기본 색상: 흰색

def apply_color_to_body(obj, color):
    """
    주어진 객체에 색상을 적용.
    """
    scaled_color = color  # 이미 0~1 범위로 가정
    obj.ViewObject.ShapeColor = scaled_color
    face_colors = [scaled_color for _ in obj.Shape.Faces]
    obj.ViewObject.DiffuseColor = face_colors
    print(f"Applied color {scaled_color} to {len(obj.Shape.Faces)} faces.")

def are_bounding_boxes_intersecting(bbox1, bbox2):
    """
    두 BoundBox가 겹치는지 확인.
    bbox1, bbox2는 FreeCAD.Base.BoundBox 객체.
    모든 축에서 겹치면 True, 아니면 False.
    """
    overlap_x = bbox1.XMin <= bbox2.XMax and bbox1.XMax >= bbox2.XMin
    overlap_y = bbox1.YMin <= bbox2.YMax and bbox1.YMax >= bbox2.YMin
    overlap_z = bbox1.ZMin <= bbox2.ZMax and bbox1.ZMax >= bbox2.ZMin
    return overlap_x and overlap_y and overlap_z

def generate_bodies(input_file):
    """
    Reads the input file and creates P-body, D-body, N-body geometries.
    Applies colors and text annotations, and displays the objects in the FreeCAD GUI.
    """
    P_bodies = []
    D_bodies = []
    N_bodies = []
    text_positions = []

    with open(input_file, "r", encoding="utf-8") as f:
        for line_number, line in enumerate(f, 1):
            cleaned_line = clean_line(line)
            if not cleaned_line.strip():
                continue

            parts = cleaned_line.strip().split(", ")
            if len(parts) < 4:
                print(f"Skipping malformed line {line_number}: {cleaned_line}")
                continue

            body_type = parts[0].upper()
            shape_type = parts[3].upper()

            # RECTANGLE 처리
            if shape_type == "RECTANGLE":
                try:
                    center_x = float(parts[4])
                    center_y = float(parts[5])
                    x_size = float(parts[6])
                    y_size = float(parts[7])
                    z_start = float(parts[1])
                    depth = float(parts[2])
                    rotation = float(parts[8]) if len(parts) > 8 else 0
                    body = create_box(center_x, center_y, x_size, y_size, z_start, depth, rotation)

                    if len(parts) == 11:
                        text = parts[10].strip('"')
                        text_z = z_start + depth
                        text_positions.append((text, (center_x-5*2/3*len(text)/2, center_y, text_z)))

                except (ValueError, IndexError) as e:
                    print(f"Error parsing RECTANGLE line {line_number}: {cleaned_line} - {e}")
                    continue

            # CIRCLE 처리
            elif shape_type == "CIRCLE":
                try:
                    center_x = float(parts[4])
                    center_y = float(parts[5])
                    radius = float(parts[6])
                    z_start = float(parts[1])
                    height = float(parts[2])
                    body = create_cylinder(center_x, center_y, radius, z_start, height)

                    if len(parts) == 9:
                        text = parts[8].strip('"')
                        text_z = z_start + height
                        text_positions.append((text, (center_x, center_y, text_z)))

                except (ValueError, IndexError) as e:
                    print(f"Error parsing CIRCLE line {line_number}: {cleaned_line} - {e}")
                    continue

            else:
                print(f"Unknown shape type on line {line_number}: {shape_type}")
                continue

            # 색상 처리
            color = parse_color(parts[9 if shape_type == "RECTANGLE" else 7])

            # P 바디 처리
            if body_type == "P":
                print(f"Processing P-body on line {line_number}")
                obj = Part.show(body)
                apply_color_to_body(obj, color)
                P_bodies.append((body, obj, color))
                # 객체 삭제
                FreeCAD.ActiveDocument.removeObject(obj.Name)

            # D 바디 처리
            elif body_type == "D":
                print(f"Processing D-body on line {line_number}")
                obj = Part.show(body)
                apply_color_to_body(obj, color)
                D_bodies.append((body, obj, color))
                # 객체 삭제
                FreeCAD.ActiveDocument.removeObject(obj.Name)

            # N 바디 처리
            elif body_type == "N":
                print(f"Processing N-body on line {line_number}")
                obj = Part.show(body)
                apply_color_to_body(obj, color)
                N_bodies.append((body, obj, color))
                # 객체 삭제
                FreeCAD.ActiveDocument.removeObject(obj.Name)

    # 상세 로그 출력
    print("=== Body generation completed ===")
    for i, (body, obj, color) in enumerate(P_bodies):
        print(f"P Body {i + 1}: Color={color}")
    for i, (body, obj, color) in enumerate(D_bodies):
        print(f"D Body {i + 1}: Color={color}")
    for i, (body, obj, color) in enumerate(N_bodies):
        print(f"N Body {i + 1}: Color={color}")

    return P_bodies, D_bodies, N_bodies, text_positions


def D_body_sub_N_body(D_body, N_SUM, D_color):
    """
    Subtract N_SUM from a single D body.
    """
    print("Processing a single D body...")
    try:
        if not are_bounding_boxes_intersecting(D_body.BoundBox, N_SUM.BoundBox):
            print("D body does not intersect with N_SUM. Skipping.")
            return D_body, None

        updated_body = D_body.cut(N_SUM)
        if updated_body.isNull():
            print("D body - N_SUM resulted in a null body. Skipping.")
            return None, None

        new_obj = Part.show(updated_body)
        apply_color_to_body(new_obj, D_color)
        return updated_body, new_obj

    except Exception as e:
        print(f"Error processing D body: {e}")
        return None, None

def fuse_P_bodies(P_bodies):
    """
    Fuse all P bodies into a single body (P_SUM).
    """
    print("Fusing P bodies into P_SUM...")
    P_SUM = None
    for i, (P_body, _, _) in enumerate(P_bodies):  # 튜플 형태에서 P_body를 언팩
        print(f"Fusing P body {i + 1} into P_SUM...")
        if P_SUM is None:
            P_SUM = P_body
        else:
            P_SUM = P_SUM.fuse(P_body)
    if P_SUM:
        debug_body(P_SUM, "P_SUM")
    else:
        print("No P bodies found.")
    return P_SUM

def fuse_N_bodies(N_bodies):
    print("Fusing N bodies into N_SUM...")
    N_SUM = None
    for i, (N_body, _, _) in enumerate(N_bodies):
        print(f"Fusing N body {i + 1} into N_SUM...")
        if N_SUM is None:
            N_SUM = N_body
        else:
            N_SUM = N_SUM.fuse(N_body)
    if N_SUM:
        debug_body(N_SUM, "N_SUM")
    return N_SUM


def main():
    # Create new document
    doc = FreeCAD.newDocument("PPT_Model")
    input_file = r"C:\\tmp_freecad\\ppt_freecad.txt"

    # Generate bodies and text positions
    P_bodies, D_bodies, N_bodies, text_positions = generate_bodies(input_file)

    # Fuse P bodies
    P_SUM = fuse_P_bodies(P_bodies) if P_bodies else None
    print(f"P_SUM created: {'Yes' if P_SUM else 'No'}")

    # Fuse N bodies
    N_SUM = fuse_N_bodies(N_bodies) if N_bodies else None
    print(f"N_SUM created: {'Yes' if N_SUM else 'No'}")

    # P_SUM 처리 (P_SUM이 있는 경우에만)
    if P_SUM:
        # N_SUM이 있을 경우에만 cut 연산 수행
        if N_SUM:
            try:
                print("Calculating P_SUM - N_SUM...")
                P_SUM_UPDATED = P_SUM.cut(N_SUM)
                if not P_SUM_UPDATED.isNull():
                    updated_obj = Part.show(P_SUM_UPDATED)
                    apply_color_to_body(updated_obj, (0.9, 0.9, 0.9))
                else:
                    print("P_SUM - N_SUM resulted in a null body. Showing original P_SUM.")
                    updated_obj = Part.show(P_SUM)
                    apply_color_to_body(updated_obj, (0.9, 0.9, 0.9))
            except Exception as e:
                print(f"Error during P_SUM - N_SUM: {e}")
                print("Showing original P_SUM.")
                updated_obj = Part.show(P_SUM)
                apply_color_to_body(updated_obj, (0.9, 0.9, 0.9))
        else:
            # N_SUM이 없는 경우 원본 P_SUM 표시
            print("N_SUM not present. Showing original P_SUM.")
            updated_obj = Part.show(P_SUM)
            apply_color_to_body(updated_obj, (0.9, 0.9, 0.9))
    else:
        print("No P_SUM present. Proceeding with D bodies only.")

    # D_bodies 처리
    if D_bodies:
        for i, (D_body, D_obj, D_color) in enumerate(D_bodies):
            try:
                # N_SUM이 있을 경우에만 cut 연산 수행
                if N_SUM:
                    print(f"Processing D_body {i + 1} - Subtracting N_SUM...")
                    updated_D_body = D_body.cut(N_SUM)
                    if not updated_D_body.isNull():
                        updated_obj = Part.show(updated_D_body)
                    else:
                        print(f"D_body {i + 1} - N_SUM resulted in a null body. Showing original D_body.")
                        updated_obj = Part.show(D_body)
                else:
                    # N_SUM이 없는 경우 원본 D_body 표시
                    print(f"N_SUM not present. Showing original D_body {i + 1}.")
                    updated_obj = Part.show(D_body)
                
                updated_obj.Label = f"D_body_{i + 1}"
                apply_color_to_body(updated_obj, D_color)
                debug_body(updated_obj.Shape, f"D_body_{i + 1}")

            except Exception as e:
                print(f"Error processing D_body {i + 1}: {e}")
                print(f"Showing original D_body {i + 1}.")
                try:
                    updated_obj = Part.show(D_body)
                    updated_obj.Label = f"D_body_{i + 1}"
                    apply_color_to_body(updated_obj, D_color)
                except Exception as e2:
                    print(f"Error showing original D_body {i + 1}: {e2}")
    else:
        print("No D bodies present.")


    # Add texts
    # 텍스트 추가
    for text, (x, y, z) in text_positions:
        # 색상 및 텍스트 분리
        if '.' in text:
            prefix, content = text.split('.', 1)
            color_map = {
                'b': (0.0, 0.0, 0.0),  # 검정
                'w': (1.0, 1.0, 1.0),  # 흰색
                'r': (1.0, 0.0, 0.0),  # 빨강
                's': (0.0, 0.0, 1.0),  # 파랑
                'y': (1.0, 1.0, 0.0),  # 노랑
                'g': (0.0, 1.0, 0.0),  # 그린
                'o': (1.0, 0.5, 0.0),  # 오렌지
                'p': (0.5, 0.0, 0.5),  # 보라
                'i': (0.0, 0.0, 0.5),  # 남색
                'c': (0.0, 1.0, 1.0),  # 시안
            }
            color = color_map.get(prefix.lower(), (1.0, 0.0, 0.0))  # 기본값: 빨강
        else:
            content = text
            color = (0, 0, 0)  # 기본값: 검장

        # 텍스트와 접두사 분리
        if '.' in text:
            prefix, content = text.split('.', 1)  # content에 텍스트 내용을 저장
        else:
            prefix, content = '', text           # content에 원래 text 저장

        # 텍스트 위치 계산 및 추가
        add_text_to_plane(
            text=content,  # 분리된 텍스트
            position=(x - FONT_SIZE * (len(content) / 20), y - FONT_SIZE / 3),
            z_position=z + 0.5,
            height=FONT_SIZE,
            color=color  # 분리된 색상
        )


    # Set up view and save
    try:
        view = FreeCADGui.ActiveDocument.ActiveView
        view.viewAxonometric()
        focus_on_all_objects()
        FreeCADGui.updateGui()
        FreeCAD.ActiveDocument.recompute()
        FreeCAD.ActiveDocument.saveAs(r"C:\tmp_freecad\PPT_Model_with_text.FCStd")
    except Exception as e:
        print(f"Error during final processing: {e}")

if __name__ == "__main__":
    main()

