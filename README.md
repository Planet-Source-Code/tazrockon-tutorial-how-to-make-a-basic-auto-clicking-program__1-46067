<div align="center">

## Tutorial: How to make a basic auto clicking program


</div>

### Description

This tutorial will teach the reader how to make a basic clicking program that will click on coordinates set by the user every 2 seconds.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[tazrockon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tazrockon.md)
**Level**          |Beginner
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tazrockon-tutorial-how-to-make-a-basic-auto-clicking-program__1-46067/archive/master.zip)





### Source Code

<div align="center"><b><font size="4" color="#FF0000">How to make an Auto Clicker<br>
 By Tazrockon</font></b></div>
<p>Before you begin this tutorial there are some things I expect you to know.
 You should have some experience with programming in Visual Basic and you should
 at least have some basic knowledge of how to use Windows API in your programs.
 I also expect you to be able to use forms, modules, and the standard Visual
 Basic objects. Using this knowledge that already exists in your head I will
 now attempt to teach you how to make a simple auto clicker that will repeatedly
 click a certain coordinate on the screen every two seconds until told to stop.</p>
<p>The first thing you need to do is open up Visual Basic and start a Standard
 EXE. On your form place two text boxes, two labels, and four command buttons.
 Position the two text boxes side by side with some space in between in the middle
 of the left side of the form. Above each text box put the two labels. Below
 the two text boxes put two command buttons. Now, on the right side of the form
 put the other two text boxes one above the other. In the label above the first
 text box type in &quot;X Pos:&quot; (without the &quot; 's). In the other label
 type &quot;Y Pos:&quot;. Clear out the text of the two text boxes and make the
 first command button below them say &quot;Lock&quot; and the other &quot;Unlock&quot;.
 Now in the first command button on the right side of the form type in &quot;Begin&quot;
 and in the one below it type &quot;End&quot;. Now your GUI is pretty much finished.</p>
<p>Here's how the program is going to work when we are finished. When the form
 loads, the first text box will contain the user's mouse's X coordinate and the
 second will contain the user's Y coordinate as they move their mouse around
 the screen. The first command button under the mouse position boxes that reads
 &quot;Lock&quot; will be used to lock the users current mouse coordinates in
 the text boxes so the user can move their mouse around without the text box's
 numbers changing. The second command button which reads &quot;Unlock&quot; will
 be used to unlock the current coordinates in the text boxes so that the user
 can see their mouse's coordinates once again as they move their mouse across
 the screen. The first command button on the right that reads &quot;Begin&quot;
 will make the program start clicking the coordinates locked by the user and
 the &quot;End&quot; button below it will make the program stop clicking.</p>
<p>Now the light reading is over and we are ready to get down to the code. Make
 a new Module and in it add the lines:</p>
<p><font color="#FF0000">Declare Function GetCursorPos&amp; Lib &quot;user32&quot;
 (lpPoint As PointAPI)<br>
 Type PointAPI<br>
 X As Long<br>
 Y As Long<br>
 End Type</font></p>
<p><font color="#0000FF">Line1 : This is needed to tell the computer the program
 wants to get the cursor position.<br>
 Line2 : This starts PointAPI.<br>
 Line3 : Sets the X coordinate variable as Long.<br>
 Line4 : Sets the Y coordinate variable as Long.<br>
 Line5 : This ends PointAPI.</font></p>
<p>Go back to your form and add a timer. Set Enabled to True and the Interval
 to 10. Double click on the timer to get to the code window so you can make the
 timer do something. In the timer sub type:</p>
<p><font color="#FF0000">Dim pos<br>
 Dim pt As PointAPI<br>
 pos = GetCursorPos(pt)<br>
 Text1.Text = pt.X<br>
 Text2.Text = pt.Y</font></p>
<p><font color="#0000FF">Line1 : This sets pos as a variable.<br>
 Line2 : This sets the variable pt to the PointAPI used in the module.<br>
 Line3 : Sets pos equal to GetCursorPos(pt). Basicly it gets the mouse coordinates
 from the PointAPI.<br>
 Line4 : Makes Text1 read out the current X position of the mouse.<br>
 Line5 : Makes Text2 read out the current Y position of the mouse.</font></p>
<p>If you have done everything correctly you should now be able to run the program.
 When it starts up it should tell the current position of your mouse as you move
 it across the screen. Try moving your mouse to the very left bottom corner of
 your computer screen and see what it says the coordinates are. Now we can add
 the ability to Lock and Unlock coordinates. You will probably be suprised at
 how easy this is to do. Go back to your form and double click on the button
 that reads &quot;Lock&quot;. In the command sub type:</p>
<p><font color="#FF0000">Timer1.Enabled = False</font></p>
<p><font color="#0000FF">Line1 : Disables Timer1 so that it will stop reading
 out the mouse coordinates.</font></p>
<p>Go back to the form and double click on the button that reads &quot;Unlock&quot;.
 In this command sub in the code window type:</p>
<p><font color="#FF0000">Timer1.Enabled = True</font></p>
<p><font color="#0000FF">Line1 : Re-enables Timer1 so that it will start showing
 the current coords again.</font></p>
<p>Now run the program. Test out clicking the Lock and Unlock buttons. Have you
 found something bad about these buttons? Chances are you have. You can not lock
 the coords of anywhere except where the Lock button is. This can easilly be
 fixed. Go back to your form and click on the Lock button once. Now change the
 caption from &quot;Lock&quot; to &quot;&amp;Lock&quot;. Notice how the button
 now reads Lock. This means that when you run your program if you press Alt and
 L on your keyboard your program will act as if you pressed the Lock button.
 This will enable you to lock coordinates anywhere on your monitor.</p>
<p>We have set up the GUI, gotten the mouse's coordinates, and made Lock and Unlock
 buttons. What are we going to do now? We are going to make the code that will
 make our program actually click on the coordinates we lock. This is the hardest
 part of the tutorial, but if you have made it this far alright and you have
 gotten the code to get the cursor position to work, then you should be able
 to achieve our new goal. Open up the module and add the following beneath End
 Type of our Point API:</p>
<p><font color="#FF0000">Declare Function SetCursorPos Lib &quot;user32&quot;
 (ByVal X As Long, ByVal Y As Long) As Long<br>
 Declare Sub mouse_event Lib &quot;user32&quot; (ByVal dwFlags As Long, ByVal
 dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)<br>
 Public Const MOUSEEVENTF_LEFTDOWN = &amp;H2<br>
 Public Const MOUSEEVENTF_LEFTUP = &amp;H4</font></p>
<p><font color="#0000FF">Line1 : This is needed to tell the computer what to use
 to set the cursor position.<br>
 Line2 : Gets different mouse events from the computer.<br>
 Line3 : Makes the mouse event left down public so we can use it in the rest
 of our module.<br>
 Line4 : Makes the mouse event left up public so we can use it in the rest of
 our module.</font></p>
<p>Now we need to turn these API's into actions that are program can do. First
 we will make the MouseMove action that we will use to, you got it, move the
 mouse. Under the last line of API that we typed add the following code:</p>
<p><font color="#FF0000">Sub MouseMove(xP As Long, yP As Long)<br>
 Dim move<br>
 move = SetCursorPos(xP, yP)<br>
 End Sub</font></p>
<p><font color="#0000FF">Line1 : This creates the MouseMove Sub and sets the variable
 xP and yP yo Long.<br>
 Line2 : Sets the variable move.<br>
 Line3 : Sets move equal to the API SetCursorPos in (xP,yP),<br>
 Line4 : Ends the Sub.</font></p>
<p>We have the code in the module to make our mouse move, but how do we incorporate
 this into our form? Now the magic will begin. Go back to your form and add a
 second timer. Set Enabled to False and make the Interval 2000 (every two seconds).
 Double click on it and add this code:</p>
<p><font color="#FF0000">Dim xP As Long<br>
 Dim yP As Long<br>
 xP = Text1.Text<br>
 yP = Text2.Text<br>
 MouseMove (xP), (yP) </font></p>
<p><font color="#0000FF">Line1 : Sets the variable xP as Long<br>
 Line2 : Sets the variable yP as Long<br>
 Line3 : Makes xP equal the X position that's in the first text box<br>
 Line4 : Makes yP equal the Y position that's in the second text box<br>
 Line5 : Uses the MouseMove sub that we put in our module to move the mouse to
 the locked coordinates (xP and yP)</font></p>
<p>Go back to your form. Double click on the button on the right side of the form
 that reads &quot;Begin&quot;. Here we will put in the code that enables our
 second timer. Type in this code:</p>
<p><font color="#FF0000">Timer2.Enabled = True</font></p>
<p><font color="#0000FF">Line1 : This enables (turns on) Timer2 that has the MouseMove
 procedure in it.</font></p>
<p>Now go back to your form and double click on the button that reads &quot;End&quot;.
 Here we will put the code that disables the second timer, which will make the
 program stop trying to move the mouse to the locked coordinates.</p>
<p><font color="#FF0000">Timer2.Enabled = False</font></p>
<p><font color="#0000FF">Line1 : This disables Timer2 and will make the program
 stop moving the mouse.</font></p>
<p>Now run the program. Position your mouse somewhere on the screen and lock its'
 position. Click the begin button and watch your mouse move to the coordinates
 you locked. Now quickly move your mouse over the Endd button and click on it.
 This should make your program stop moving the mouse. All we have left to do
 is make the program click on the coordinate. Go back to the module and add the
 following code below our MouseMove sub:</p>
<p><font color="#FF0000">Sub LeftClick(xP As Long, yP As Long)<br>
 mouse_event MOUSEEVENTF_LEFTDOWN, xP, yP, 0, 0<br>
 mouse_event MOUSEEVENTF_LEFTUP, xP, yP, 0, 0<br>
 End Sub</font></p>
<p><font color="#0000FF">Line1 : This creates the LeftClick sub and sets the variables
 xP and yP as Long.<br>
 Line2 : Tells our program to push down the left mouse button on the coordinates.<br>
 Line3 : Tells our program to let up the left mouse button on the coordinates.<br>
 Line4 : Ends the LeftClick sub.</font></p>
<p>Go back to the form and look at the Timer2 code. Below the MouseMove line we
 need to add the code that will pull the LeftClick procedure from the module.
 This is very easy. Add this code to the Timer2 sub:</p>
<p><font color="#FF0000">LeftClick (xP), (yP)</font></p>
<p><font color="#0000FF">Line1 : Uses the LeftClick sub that we put in our module
 to left click the coordinates.</font></p>
<p>We should now have a fully functional automatic clicking program that will
 click on given coordinates every two seconds until told to stop. Some things
 to try are:</p>
<p><font color="#009900">*Make a program that clicks on 2 or more points.<br>
 *Use a slider to allow the user to change the clicking interval.<br>
 *Make a mouse macro that will run Paint by clicking on Start, moving the mouse
 up to Programs, move the mouse up to Accessories, and clicking on Paint. </font></p>

