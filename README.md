# GoogleScriptUngu

GoogleScriptUngu

## Face Recognition (Python)

This repository now includes a simple Python face recognition demo using a webcam.

### Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Prepare known faces

Create a directory named `known_faces` with one subfolder per person. Put JPG images for each person inside their folder:

```
known_faces/
  Alice/
    alice1.jpg
  Bob/
    bob1.jpg
```

### Run

```bash
python face_recognition_app.py --known-dir known_faces
```

Press `q` to quit the app.
