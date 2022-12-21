from PIL import Image
import os


def png2coe(imageName):
    """
    Transfer from png to coe
    :imageName: TODO
    :returns: TODO
    """
    im = Image.open(imageName)
    width = im.size[0]
    height = im.size[1]
    forward, end = os.path.splitext(imageName)
    outfile = forward + ".coe"
    with open(outfile, 'w') as f:
        f.write('MEMORY_INITIALIZATION_RADIX=2;\n')
        f.write('MEMORY_INITIALIZATION_VECTOR=\n')
        for h in range(height):
            for w in range(width):
                pixel = im.getpixel((w, h))  # get pixel
                # 将24位RGB转化为12位RGB
                r = pixel[0] >> 4
                g = pixel[1] >> 4
                b = pixel[2] >> 4
                f.write(
                    bin(r)[2:].zfill(4) + bin(g)[2:].zfill(4) +
                    bin(b)[2:].zfill(4))
                if w == width - 1 and h == height - 1:
                    f.write(';')
                else:
                    f.write(',\n')
    f.close()


if __name__ == '__main__':
    ImageName = 'background.png'
    png2coe(ImageName)
