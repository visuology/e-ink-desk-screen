import os
from importlib import import_module

from PIL import Image, ImageDraw, ImageFont


driver_name = os.environ.get("DESK_EPD_DRIVER", "epd2in7_V2")
driver = import_module(f"waveshare_epd.{driver_name}")

epd = driver.EPD()
epd.init()
epd.Clear()

canvas = Image.new("1", (epd.height, epd.width), 255)
draw = ImageDraw.Draw(canvas)
font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 24)
draw.rectangle((0, 0, epd.height, 42), fill=0)
draw.text((10, 8), "Desk Display", font=font, fill=255)
draw.text((10, 60), driver_name, font=font, fill=0)

rotation = int(os.environ.get("DESK_DISPLAY_ROTATION", "0"))
epd.display(epd.getbuffer(canvas.rotate(rotation) if rotation else canvas))
epd.sleep()

