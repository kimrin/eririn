# eririn

Generate Excel Graph Paper sheets from image files.

motto: to create best quality of eririn.xlsx.

## What is eririn?

Eririn is a simpile python script that generates Excel WorkBook file (.xlsx) from image file(s).

## What is HouganShi-Excel(方眼紙Excel) = Graph Paper Excel file

* In Japan, there is a silly old customs of a way to create Excel file that each character in each cells

* It is like a appearance of Graph Paper, so usually we called it "HouganShi-Excel" (HouganShi=Graph Paper in Japanese.)

## Technical Descriptions:

* kimrin starts this pip module as "Parody" program, but faced Excel file color limitations (=65536 colors in a single WorkBook file).
* so I implemented a color reduction algorithm by using K-means clustering......

### Strategy of conversions:

* If file size = (width, height) is small and total number of color is smaller than 65536, then simply create original image
dot-by-dot Excel Sheet (This takes about 20 seconds.)

* If colors exceeds 65536, we attempt to reduce colors within 65536 that keeps original image sile (width, height)
(This takes about 30 minutes(!))

* in another cases, simply resize picture and attempts generation of the Excel WorkBook.

* kimrin is not yet implemented multipile image files conversion under the restriction of 65536 colors.
* so the program can create multiple seets in each image files, but will fail in open by the MicroSoft Excel...

## Who is eririn?

[https://en.wikipedia.org/wiki/Eriko_Yamaguchi](https://en.wikipedia.org/wiki/Eriko_Yamaguchi)

(Woman Japanese Chess professional: and her nickname is "eririn".)

## How to use:

### Install

Simply using pip.

```bash
pip install eririn
```

### run
pip installs eririn command, so simply invoke this command from CLI.

```bash
eririn eririn.xlsx img/eririn.png
```

#### take care

* xlsx file name is first argument of the command.
* Usualy a image file is specified. multiple files can be specified, but likely occuors exceeding of color pallete.

### LICENCE: MIT
