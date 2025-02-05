import os
import time


def wait_for_downloads_to_complete(download_dir, timeout=1500):
    """
    Waits until all downloads are complete (i.e., no .crdownload files in the directory).
    :param download_dir: The directory where files are being downloaded.
    :param timeout: Maximum time to wait for downloads to complete (in seconds).
    """
    end_time = time.time() + timeout
    while time.time() < end_time:
        # Check if any file with .crdownload exists in the download directory
        if not any(f.endswith('.crdownload') for f in os.listdir(download_dir)):
            return True
        time.sleep(1)
    raise TimeoutError("Timed out waiting for downloads to complete.")