import logging
from PIL import Image, ImageGrab
from .models import ScreenShotRegion, ScreenshotMode


class ScreenShot:

    def __init__(self, mode: str, region=ScreenShotRegion):
        logging.log(logging.INFO, f"Initializing Screenshot: {mode}, Region: {region}")
        self.mode = mode
        self.region = region
        self.ss_image = None

    def take_screenshot(self):

        match self.mode:
            case ScreenshotMode.FULLSCREEN:
                try:
                    logging.log(logging.INFO, f"Grabbing {self.mode} screenshot")
                    self.ss_image = ImageGrab.grab(bbox=False)
                    logging.log(logging.INFO, f"Grabbed {self.mode} screenshot")
                except Exception as ex:
                    logging.log(
                        logging.ERROR,
                        f"Grabbing Screenshot mode: {self.mode}, Screenshot failed: {ex}",
                    )
            case ScreenshotMode.PARTIAL:
                try:
                    logging.log(logging.INFO, f"Grabbing {self.mode} screenshot")
                    self.ss_image = ImageGrab.grab(
                        bbox=self.region.get_region(),
                    )
                    logging.log(logging.INFO, f"Grabbed {self.mode} screenshot")
                except Exception as ex:
                    logging.log(
                        logging.ERROR,
                        f"Grabbing Screenshot mode: {self.mode}, Screenshot failed: {ex}",
                    )
            case ScreenshotMode.MULTIDISPLAY:
                try:
                    logging.log(logging.INFO, f"Grabbing {self.mode} screenshot")
                    self.ss_image = ImageGrab.grab(all_screens=True)
                    logging.log(logging.INFO, f"Grabbed {self.mode} screenshot")
                except Exception as ex:
                    logging.log(
                        logging.ERROR,
                        f"Grabbing Screenshot mode: {self.mode}, Screenshot failed: {ex}",
                    )

            case ScreenshotMode.WINDOW:
                try:
                    logging.log(logging.INFO, f"Grabbing {self.mode} screenshot")
                    self.ss_image = ImageGrab.grab(
                        bbox=self.region.get_region(),
                    )
                    logging.log(logging.INFO, f"Grabbed {self.mode} screenshot")
                except Exception as ex:
                    logging.log(
                        logging.ERROR,
                        f"Grabbing Screenshot mode: {self.mode}, Screenshot failed: {ex}",
                    )

    def get_screenshot(self) -> Image.Image:
        logging.log(logging.INFO, f"Taking Screenshot")
        self.take_screenshot()

        if self.ss_image:
            logging.log(logging.INFO, f"Screenshot image taken and returning")
            return self.ss_image

        logging.log(logging.INFO, f"Error in Taking screenshot")
        exit(-1)

    def save_screenshot(self):
        logging.log(logging.INFO, "Saving Screenshot")
        # Screenshot FULL Image PATH
        save_image_path = ""
        if self.ss_image:
            self.ss_image.save("")
            logging.log(logging.INFO, f"Saved Screenshot at: {save_image_path}")

        logging.log(logging.ERROR, "Screenshot NULL")
