import re
from googleapiclient.http import MediaIoBaseDownload
import io

def descargarImagenDrive(drive_service, file_id: str) -> io.BytesIO:
    """
    Usa drive_service (ya autenticado) para descargar file_id y retorna BytesIO con los bytes.
    """
    request = drive_service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    buffer.seek(0)
    return buffer



def parseFileId(url: str) -> str:
    """
    Extrae el file_id de un link con formato:
      https://drive.google.com/open?id=<FILE_ID>
    Retorna None si no encuentra nada.
    """
    m = re.search(r'[?&]id=([^&]+)', url)
    if m:
        return m.group(1)
    return None