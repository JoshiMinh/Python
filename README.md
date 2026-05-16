# Python

A collection of Python scripts organized by library and purpose.

## Structure

```
Python/
├── manim/
│   ├── 2d/          # 2D animations (scenes, graphs, shapes)
│   └── 3d/          # 3D animations (surfaces, spheres)
│
├── turtle/
│   ├── art/         # Drawings: spirographs, landscapes, patterns
│   ├── drawings/    # Basic shapes (oval, star, triangle, squares)
│   └── flags/       # Country flag recreations
│
├── tkinter/         # Desktop GUI apps (calculator, notepad, image viewer)
│
├── standalone/
│   ├── games/       # TicTacToe
│   ├── bots/        # TelegramBot
│   ├── tools/       # PasswordGenerator, SpeedTest, URLShortener, etc.
│   ├── data/        # MaziiCSVConvert, PyMySQL
│   └── media/       # SpeechToText, FacialRecognition
│
└── requirements/
    ├── base.txt      # Shared deps (numpy, sympy, Pillow, colorama)
    ├── manim.txt     # Manim + base
    └── standalone.txt # All standalone deps + base
```

## Install

```bash
# Install only what you need:
pip install -r requirements/manim.txt
pip install -r requirements/standalone.txt

# Or just shared base deps:
pip install -r requirements/base.txt
```

> `turtle` and `tkinter` use Python's standard library — no install needed.
