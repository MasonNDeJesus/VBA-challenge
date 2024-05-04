# VBA-challenge
Dear lord this one kicked my but. Ran into so many errors, and had to use On Error Resume Next in order to get around them, basically bulldozing through the erros. As such, the results may not be entirely accurate as Excel 2010 seems to have problems with over flow
Also LongLong is not avilable for 2010 VBA (even though that's when it was supposed to have been added). As such, I used LongPtr, which got around some issues, though I still ran into the overflow error til I used my work around.
Ran into an issue with returhning the total volume I think, but I don't know if that error will pop up again. I don't really know where or why it occured but I hope it doesn't print the same number multiple times like it kept doing again.
