from pypdf import PdfReader, PdfWriter


def merge_first_pages_and_save(files, output_path):
    writer = PdfWriter()

    for file in files:
        try:
            reader = PdfReader(file)
        except FileNotFoundError:
            # 파일이 없으면 스킵
            print(f"Warning: 파일을 찾을 수 없습니다: {file}")
            continue

        if not reader.pages:
            # 페이지가 하나도 없으면 스킵
            print(f"Warning: 페이지가 없습니다: {file}")
            continue

        # 첫 페이지만 추가
        writer.add_page(reader.pages[0])

    # 결과 쓰기
    with open(output_path, "wb") as out_f:
        writer.write(out_f)
