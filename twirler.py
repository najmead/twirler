from PIL import Image, ImageOps

cover = Image.open("P00001.jpg")
cover = cover.resize((cover.size[0]*3, cover.size[1]*3), Image.ANTIALIAS)
cover = ImageOps.expand(cover, border=5, fill="black")
cover = cover.convert("RGBA")
print('Saving cover')
cover.save('cover.png')
cwidth, cheight = cover.size

background = Image.new("RGBA", (cwidth,cheight), color="white")
background = ImageOps.expand(background, border=5, fill="black")
background = background.rotate(1.5, expand=1)
bwidth, bheight = background.size

print("Saving background")
background.save("Foreground.png")

background.paste(cover, (int((bwidth-cwidth/2)),int((bheight-cheight)/2)))

print("Saving merged")
background.save("Merged.png")




