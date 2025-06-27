import argparse
import cv2


def stream_ipcam(url: str) -> None:
    """Display frames from an IP camera using OpenCV.

    Press ESC to exit.
    """
    cap = cv2.VideoCapture(url)
    if not cap.isOpened():
        raise RuntimeError(f"Unable to open stream: {url}")

    while True:
        ret, frame = cap.read()
        if not ret:
            break
        cv2.imshow('IP Camera Stream', frame)
        if cv2.waitKey(1) == 27:  # ESC key
            break

    cap.release()
    cv2.destroyAllWindows()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Stream an IP camera using OpenCV')
    parser.add_argument('url', help='RTSP/HTTP URL of the camera')
    args = parser.parse_args()
    stream_ipcam(args.url)
