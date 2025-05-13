from enum import Enum, auto


class ScreenshotMode(Enum):
    FULLSCREEN = auto()
    WINDOW = auto()
    MULTIDISPLAY = auto()
    PARTIAL = auto()


class ScreenShotRegion:

    def __init__(self, left=0, top=0, right=1920, bottom=1080):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom

    def set_left(self, left: int) -> None:
        self.left = left

    def set_right(self, right: int) -> None:
        self.right = right

    def set_top(self, top: int) -> None:
        self.top = top

    def set_bottom(self, bottom: int) -> None:
        self.bottom = bottom

    def get_region(self):
        return (self.left, self.top, self.right, self.bottom)
