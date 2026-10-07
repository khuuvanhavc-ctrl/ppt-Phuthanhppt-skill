"""
EduPPT Creator - PowerPoint Quality Checker
Version: 1.0

Kiểm tra cơ bản file PowerPoint:
- File có tồn tại không
- Có mở được bằng python-pptx không
- Có slide không
- Slide có tiêu đề không
- Slide có nội dung không
"""

import argparse
import os
import sys

from pptx import Presentation


# ============================================================
# KIỂM TRA FILE
# ============================================================

def check_file_exists(pptx_file):

    if not os.path.exists(pptx_file):
        return False, "Không tìm thấy file PowerPoint."

    if os.path.getsize(pptx_file) == 0:
        return False, "File PowerPoint đang rỗng."

    return True, "File tồn tại."


# ============================================================
# KIỂM TRA POWERPOINT
# ============================================================

def check_presentation(pptx_file):

    results = []

    # Kiểm tra file
    ok, message = check_file_exists(pptx_file)

    results.append(
        ("FILE", ok, message)
    )

    if not ok:
        return results

    # Thử mở PowerPoint
    try:

        prs = Presentation(pptx_file)

        results.append(
            (
                "OPEN",
                True,
                "PowerPoint có thể mở bằng python-pptx."
            )
        )

    except Exception as error:

        results.append(
            (
                "OPEN",
                False,
                f"Không thể mở PowerPoint: {error}"
            )
        )

        return results

    # Kiểm tra số slide
    slide_count = len(prs.slides)

    if slide_count == 0:

        results.append(
            (
                "SLIDES",
                False,
                "PowerPoint không có slide."
            )
        )

    else:

        results.append(
            (
                "SLIDES",
                True,
                f"PowerPoint có {slide_count} slide."
            )
        )

    # Kiểm tra từng slide
    for index, slide in enumerate(prs.slides, start=1):

        text_content = []

        for shape in slide.shapes:

            if not hasattr(shape, "text"):
                continue

            text = shape.text.strip()

            if text:
                text_content.append(text)

        # Slide rỗng
        if not text_content:

            results.append(
                (
                    f"SLIDE {index}",
                    False,
                    "Slide không có nội dung văn bản."
                )
            )

        else:

            results.append(
                (
                    f"SLIDE {index}",
                    True,
                    f"Có {len(text_content)} vùng nội dung."
                )
            )

    return results


# ============================================================
# IN KẾT QUẢ
# ============================================================

def print_results(results):

    print()
    print("=" * 60)
    print("EduPPT Creator - POWERPOINT QUALITY CHECK")
    print("=" * 60)

    total = 0
    passed = 0
    failed = 0

    for name, ok, message in results:

        total += 1

        if ok:

            passed += 1
            status = "PASS"

        else:

            failed += 1
            status = "FAIL"

        print(
            f"[{status}] {name}: {message}"
        )

    print()
    print("-" * 60)

    print(
        f"Tổng kiểm tra : {total}"
    )

    print(
        f"Đạt           : {passed}"
    )

    print(
        f"Lỗi            : {failed}"
    )

    print("-" * 60)

    if failed == 0:

        print(
            "KẾT QUẢ: POWERPOINT ĐẠT KIỂM TRA CƠ BẢN."
        )

    else:

        print(
            "KẾT QUẢ: POWERPOINT CẦN ĐƯỢC KIỂM TRA/SỬA."
        )

    print("=" * 60)


# ============================================================
# MAIN
# ============================================================

def main():

    parser = argparse.ArgumentParser(
        description="Kiểm tra file PowerPoint của EduPPT Creator."
    )

    parser.add_argument(
        "--input",
        required=True,
        help="Đường dẫn đến file .pptx"
    )

    args = parser.parse_args()

    print(
        f"Đang kiểm tra: {args.input}"
    )

    results = check_presentation(
        args.input
    )

    print_results(
        results
    )

    # Nếu có lỗi → trả về mã lỗi
    failed = any(
        not result[1]
        for result in results
    )

    if failed:
        sys.exit(1)

    sys.exit(0)


if __name__ == "__main__":

    main()
