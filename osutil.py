import os


def get_hwp_and_hwpx_files_under(folder_path: str) -> list[str]:
    """
    Get all hwp and hwpx files under the given folder path.

    Args:
        folder_path (str): The path to the folder.

    Returns:
        list[str]: A list of hwp and hwpx file paths.
    """
    hwp_files = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.hwp') or file.endswith('.hwpx'):
                hwp_files.append(os.path.join(root, file))
    return hwp_files
