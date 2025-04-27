import sys
import os
import hwputil
import pdfutil

if __name__ == "__main__":
    base_folder_path = sys.argv[1]

    try:
        target_files = map(lambda x: os.path.join(base_folder_path, x), filter(lambda x: x.endswith(
            '.hwpx') or x.endswith('.hwp'), os.listdir(base_folder_path)))

        pdfs = hwputil.generate_pdfs(target_files)
        pdfutil.merge_first_pages_and_save(
            pdfs, os.path.join(base_folder_path, '합본.pdf'))
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)
    finally:
        hwputil.HWPX.Quit()
