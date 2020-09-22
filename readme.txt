
I got the majority of this work from a comment that Ace315 left with a similar program that required a .dll to work.

This program demonstrates how to create a shaped form or 'skin' using any picture as a base.
To test the program, run it.  You can drag the "Hello", double-click to end the program.
To use it with another form, or picture do as follows:

- Cut & paste the contents of Form_load() into your new form
- Set the form's border style to 'none'
- Put the desired picture into the form's 'picture' property
- Set the 'transparent' color in the GetBitmapRegion function's second parameter.
        Currently white is transparent, but you can set it to anything.

Good Luck.

