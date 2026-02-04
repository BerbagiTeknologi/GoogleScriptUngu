"""Simple face recognition app using a webcam.

Usage:
  python face_recognition_app.py --known-dir known_faces

Directory structure:
  known_faces/
    Alice/
      alice1.jpg
      alice2.jpg
    Bob/
      bob1.jpg
"""

from __future__ import annotations

import argparse
import pathlib
from dataclasses import dataclass
from typing import Dict, List, Tuple

import cv2
import face_recognition


@dataclass(frozen=True)
class KnownFace:
    name: str
    encoding: List[float]


def load_known_faces(known_dir: pathlib.Path) -> List[KnownFace]:
    known_faces: List[KnownFace] = []
    if not known_dir.exists():
        raise FileNotFoundError(f"Known faces directory not found: {known_dir}")

    for person_dir in sorted(p for p in known_dir.iterdir() if p.is_dir()):
        for image_path in sorted(person_dir.glob("*.jpg")):
            image = face_recognition.load_image_file(str(image_path))
            encodings = face_recognition.face_encodings(image)
            if not encodings:
                continue
            known_faces.append(KnownFace(name=person_dir.name, encoding=encodings[0]))

    if not known_faces:
        raise ValueError(
            "No encodings found. Ensure there are JPG images with clear faces."
        )
    return known_faces


def recognize_faces(
    frame_rgb, known_faces: List[KnownFace], tolerance: float
) -> Tuple[List[Tuple[int, int, int, int]], List[str]]:
    face_locations = face_recognition.face_locations(frame_rgb)
    face_encodings = face_recognition.face_encodings(frame_rgb, face_locations)

    known_encodings = [kf.encoding for kf in known_faces]
    names: List[str] = []

    for encoding in face_encodings:
        matches = face_recognition.compare_faces(
            known_encodings, encoding, tolerance=tolerance
        )
        name = "Unknown"
        if True in matches:
            matched_indices = [i for i, matched in enumerate(matches) if matched]
            votes: Dict[str, int] = {}
            for idx in matched_indices:
                person_name = known_faces[idx].name
                votes[person_name] = votes.get(person_name, 0) + 1
            name = max(votes, key=votes.get)
        names.append(name)

    return face_locations, names


def draw_labels(frame_bgr, face_locations, names):
    for (top, right, bottom, left), name in zip(face_locations, names):
        cv2.rectangle(frame_bgr, (left, top), (right, bottom), (0, 255, 0), 2)
        cv2.rectangle(
            frame_bgr, (left, bottom - 35), (right, bottom), (0, 255, 0), cv2.FILLED
        )
        cv2.putText(
            frame_bgr,
            name,
            (left + 6, bottom - 6),
            cv2.FONT_HERSHEY_DUPLEX,
            0.8,
            (0, 0, 0),
            1,
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Webcam face recognition demo")
    parser.add_argument(
        "--known-dir",
        type=pathlib.Path,
        default=pathlib.Path("known_faces"),
        help="Directory with subfolders of known faces",
    )
    parser.add_argument(
        "--camera",
        type=int,
        default=0,
        help="Camera index for cv2.VideoCapture",
    )
    parser.add_argument(
        "--tolerance",
        type=float,
        default=0.5,
        help="Face comparison tolerance (lower is stricter)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    known_faces = load_known_faces(args.known_dir)

    video_capture = cv2.VideoCapture(args.camera)
    if not video_capture.isOpened():
        raise RuntimeError("Unable to access the camera.")

    try:
        while True:
            ret, frame_bgr = video_capture.read()
            if not ret:
                break

            frame_rgb = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)
            face_locations, names = recognize_faces(
                frame_rgb, known_faces, tolerance=args.tolerance
            )
            draw_labels(frame_bgr, face_locations, names)

            cv2.imshow("Face Recognition", frame_bgr)
            if cv2.waitKey(1) & 0xFF == ord("q"):
                break
    finally:
        video_capture.release()
        cv2.destroyAllWindows()


if __name__ == "__main__":
    main()
