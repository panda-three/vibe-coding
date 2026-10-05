import sys, os
from PIL import Image, ImageDraw, ImageFont

# args: pages_dir start end title out_path
pages_dir = sys.argv[1]
start = int(sys.argv[2])
end = int(sys.argv[3])
title = sys.argv[4]
out = sys.argv[5]

cols = 4
n = end - start + 1
rows = (n + cols - 1) // cols

cell_w = 900
pad = 36
label_h = 56
header_h = 120

imgs = []
for i in range(start, end + 1):
    p = os.path.join(pages_dir, f"p{i:03d}.png")
    im = Image.open(p).convert("RGB")
    h = int(im.height * cell_w / im.width)
    imgs.append(im.resize((cell_w, h), Image.LANCZOS))

img_h = max(im.height for im in imgs)
cell_h = label_h + img_h
W = cols * cell_w + (cols + 1) * pad
H = header_h + rows * cell_h + (rows + 1) * pad

canvas = Image.new("RGB", (W, H), "white")
draw = ImageDraw.Draw(canvas)

fp = "/System/Library/Fonts/PingFang.ttc"
title_font = ImageFont.truetype(fp, 46)
label_font = ImageFont.truetype(fp, 32)

# header bar
draw.rectangle([0, 0, W, header_h - 20], fill=(238, 245, 253))
draw.text((pad, 34), title, fill=(0, 82, 204), font=title_font)

for idx, im in enumerate(imgs):
    r = idx // cols
    c = idx % cols
    x = pad + c * (cell_w + pad)
    y = header_h + pad + r * (cell_h + pad)
    page = start + idx
    draw.rectangle([x, y, x + cell_w, y + label_h], fill=(224, 238, 253))
    draw.text((x + 14, y + 8), f"P{page:02d}", fill=(0, 82, 204), font=label_font)
    canvas.paste(im, (x, y + label_h))
    draw.rectangle([x, y + label_h, x + cell_w, y + label_h + im.height],
                   outline=(205, 205, 205), width=2)

canvas.save(out, quality=92)
print("saved", out, canvas.size)
