<div align="center">

## Pixels to Twips


</div>

### Description

This converts pixels to twips
 
### More Info
 
# of Pixels or twips

This converts pixels to twips, if you get the rect of a window it's in pixels, to make your form that size you need to convert the pixels to twips.

As someone pointed out in one of my other submisions, it might be a little faster to take away the function and just use the insides where needed but i find the function saves some typing and makes the code easier to read.

# of pixels or twips... see code


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Unknown
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pixels-to-twips__1-2110/archive/master.zip)





### Source Code

```
Function PixelsToTwips_height(pxls)
PixelsToTwips_height = pxls * screen.TwipsPerPixelY
end function
Function PixelsToTwips_width(pxls)
PixelsToTwips_width = pxls * screen.TwipsPerPixelX
end function
'This next part reverses the las although you should
'be able to use basic math
Function TwipsToPixels_height(pxls)
PixelsToTwips_height = pxls \ screen.TwipsPerPixelY
end function
Function TwipsToPixels_width(pxls)
PixelsToTwips_width = pxls \ screen.TwipsPerPixelX
end function
```

