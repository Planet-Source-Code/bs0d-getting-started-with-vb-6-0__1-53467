<div align="center">

## Getting Started with VB 6\.0


</div>

### Description

Comprehensive tutorial that is aimed at the beginning programmer that will go through all of the basics in learning Visual Basic 6.0.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[bs0d](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bs0d.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bs0d-getting-started-with-vb-6-0__1-53467/archive/master.zip)





### Source Code

<font size="2"><h4>-Getting Started-</h4>
<br>
<p>Visual Basic is an object-oriented programming language that uses the Microsoft Windows platform. The programs that are created using Visual Basic will look and act like standard Windows programs. Visual Basic provides one the tools to create windows with elements such as menus, text boxes, command buttons, option buttons, list boxes and scroll bars.</p>
<p>This tutorial does not completely cover all ‘basis’ of Visual Basic. This is just basically an over view to learn a bit about VB. Databases, Crystal Reports and others are not included to keep this tutorial under 500 pages and to keep carpal tunnel from occurring.</p>
<p>This tutorial was written assuming that you possibly have some programming knowledge. Like what an IF / THEN / ELSE statement is. So, if you have taken Basic Or QBasic, that’s ideal before reading this tutorial.</p>
<p>Also, it would not help much reading this tutorial if you do not have VB 6.0 on your computer. I think you can download it from Kazaa Lite (kinda big though) if you do not have it. But get it if you don’t have it, open it and read this.</p>
<br>
<h4>Procedural vs. Non-Procedural Languages</h4>
<br>
<p><u>Procedural Languages</u> - Programming languages that have a set plan that you follow to execute a program. You have a series of statements that you execute where you start with the first statement. The statements are executed in order from beginning to end. The program terminates after the last statement is executed.</p>
<br>
<i>Examples of Procedural Languages:</i>
<br>
<ul>
<li>FORTRAN</li>
<li>COBOL</li>
<li>BASIC</li>
<li>C</li>
<li>PASCAL</li>
</ul>
<br>
and more... <br>
<p><u>Non-Procedural Languages</u> - Object-Oriented programming languages that are Event-driven. Here you don’t just have a series of statements that are executed, you have several choices of different things you can do in a program. You select the event that you want to occur. Only the code for that event is executed.</p>
<br>
<i>Examples of Procedural Languages:</i>
<br>
<ul>
<li>Visual Basic</li>
<li>C++</li>
<li>DELPHI</li>
<li>JAVA</li>
</ul>
<h5>Labels, Text Boxes and Command Buttons</h5>
<br>
<p><B>Label</b> - A control that is used to display text as a caption. Labels cannot be altered by the used. They are simply used to display headings and results of processing as well as other information on a form. Labels begin with a .lbl extension.</p>
<br>
<i>Examples of labels:</i>
<br>
<ul>
<li>lblName</li>
<li>lblAddress</li>
<li>lblCity</li>
</ul>
<br>
<p><b>Text Box</b> - Control that is used to enter information onto a form. Text boxes can have focus and are the primary means of providing input to a project in Visual Basic. Text boxes begin with the prefix of .txt</p>
<br>
<i>Examples of Text Boxes:</i>
<br>
<ul>
<li>txtSponge</li>
<li>txtBob</li>
<li>txtNum1</li>
</ul>
<br>
<p><b>Command Buttons</b> - Control used to activate a procedure. When a command button is clicked on or activated using an access key, and event will take place (the code behind the command button is executed.) Command buttons begin with the prefix .cmd</p>
<br>
<h5>Types of Errors</h5>
<br>
<ul>
<li><u>Syntax Errors</u> - Compiler errors that occur when your project code is converted to machine language. These are errors that occur when you violate the syntax rules of Visual Basic. VB will highlight these errors in RED.</li>
<li><u>Run-Time Errors</u> - Errors that occur when programming is actually executing. These types of errors will cause the program to terminate abnormally. Examples of run-time errors are division by zero, or trying to do calculations with non-numeric data. VB will display a dialog box, pull up the affected code on screen, and highlight where the error occurred in YELLOW.</li>
<li><u>Logical Errors</u> - Errors which allow your project to run, but will produce incorrect results. These errors are tough to find and debug because no error message will be displayed or highlighted. You will have to dig through the code to find the error(s). Examples would be : Adding 3 instead of subtracting 3, or displaying the wrong info in a label.</li>
</ul>
<br>
<h5>Special Functions and Methods</h5>
<br>
<p><b>VAL()</b> - A function that will convert a string to a numeric value. It begins with the left-most character of the string. If that character is a numeric digit, decimal point or a sign, VAL will convert the character to a numeric and move on to the next character. Once a non-numeric character is found, the VAL function will stop. </p>
<p><b>PrintForm</b> - A method that will print the current form on the printer. This is executed during run-time. </p>
<p><b>FORMAT</b> - This is used to format a variable. Variables can be formatted as fixed numbers, currency and percents. </p>
<br>
Syntax :<br>
<pre>
< var > = FORMAT( < var > ," < type of format > ")
</pre>
<br>
Examples:<br>
<pre>
lblTotal = Format(lblTotal,"Currency") 'Formats number with a $ and two decimal places.
lblArea = Format(lblTotal,"Fixed") ' Formats number as decimal with two decimal places.
LblPctTotal = Format(lblPctTotal,"Percent") 'Formats number as percent with two decimal places and adds % to end of number.
</pre>
<br>
<p><b>FormatCurrency</b> - This function will take a number and format it as currency with two decimal places and a $ sign. There is only one argument, the number you want to format.</p>
<br>
<pre>
lblSum = FormatCurrency(NumToFormat)
lblTotalAmount = FormatCurrency(curTotal)
</pre>
<br>
<p><b>FormatNumber</b> - This function will format a number as a decimal with a set number of decimal places. The first argument is the number you want to format. The second is the number of decimal places you want the number to have.
</p>
<br>
<pre>
lblTotal = FormatNumber(NumToFormat,NumOfDecimalPlaces)
lblSum = FormatNumber(Sum,2)
</pre>
<br>
<p><b>FormatPercent</b> - This function will take the number and format it as a percent. The function will multiply by 100 and add a percent sign to the end of the number. The number by default will be rounded to zero decimal places unless you specify a 2nd argument.</p>
<br>
<pre>
lblTotalPct = FormatPercent(NumToFormatAsPercent)
lblSumPct = FormatPercent(NumToFormatAsPercent, NumOfPlaces)
lblCurrInterest = FormatPercent(CurrentInterest)
</pre>
<br>
<p><b>FormatDate and Time</b> - This function will take a string and format it as a Date, Time or both. The first argument is the date or time to be formatted. The 2nd argument is the format of date/time you wish to use.</p>
<br>
<pre>
lblCurrentDate = FormatDateTime(StartingDate, vbShortDate)
lblCurrentDate = FormatDateTime(EndingDate, vbLongDate)
lblCurrentTime = FormatDateTime(StartingTime, vbShortTime)
lblCurrentTime = FormatDateTime(StartingTime, vbLongTime)
</pre>
<h4>Variables and Constants</h4>
<br>
<p><b>Option Explicit</b> - A statement that will force you to declare all variables used in your program. If you fail to declare a variable, an error will occur when you run your program.</p>
<br>
<h5>DECLARING A VARIABLE:</h5>
<br>
<p>When you declare a variable, VB reserves space in the computers memory and assigns it a name. When declaring variables, there are certain rules that must be followed.</p>
<br>
<b>RULES (for variable names):</b>
<br>
<ul>
<li>Variables must begin with a letter.</li>
<li>Can consist of letters, digits and the underscore.</li>
<li>Cannot contain any spaces or periods.</li>
<li>Must be 1 to 255 positions in length</li>
<li>Cannot use Reserved Words as variable names.</li>
</ul>
<br>
<i><u>NOTE:</u> Rules ( above ) also apply to naming controls that are placed on a form.</i>
<br>
The DIM statement is used to declare all variables in VB.
<br>
<i>Syntax:</i>
<br>
<b>DIM < variable_name > as < datatype > </b>
<br>
<u>Examples:</u>
<br>
<pre>
DIM inNum1 as Integer
DIM strName as String
DIM curTotalAmount as Currency
</pre>
<br>
<h4>DECLARING A CONSTANT:</h4>
<br>
<p>Constants are always declared using the keyword CONST. You give the constant a name, data type and value. Once a value is declared a constant, it can never be changed again in the program. Trying to change a constant will cause an error. The rules for variables also apply to constants.</p>
<br>
<i>Syntax:</i>
<br>
<b>DIM < constant name > as < datatype > = < value > </b>
<br>
<u>Examples:</u>
<br>
<pre>
DIM sngPI as Single = 3.14
DIM curTaxRate as Currency = .07
</pre>
<br>
<p><b>NAMED CONSTANTS</b> - These are constants that you name yourself with the CONST keyword. Named constants can come in the form as numeric constants and string constants.</p>
<p><b>NUMERIC CONSTANTS </b>- Constants that can contain only the digits 0-9, a decimal pt and sign. </p>
<p><b>STRING CONSTANTS</b> - Constants that can contain letters, digits, and special characters such as @#$%^&*. String constants must be enclosed in double quotes.</p>
<p><b>INSTRINCT CONSTANTS</b> - System-Defined constants that are built into VB.</p>
<br>
<u>Examples of Instrinct Constants:</u>
<br>
<ul>
<li>vbRed</li>
<li>vbGreen </li>
<li>vbBlue</li>
<li>Checked</li>
<li>vbYellow</li>
<li>UnChecked</li>
</ul>
<br>
<h4>SCOPES OF VARIABLES:</h4>
<br>
<b>Scope</b> - Is a term used to refer to the visibility of a variable.
<br>
<b>Lifetime</b> - The period of time that variables exist.
<br>
<h5>3 LEVELS OF SCOPE:</h5>
<br>
<ul>
<li><u>Global Variable</u> - Variable accessible anywhere in VB, in all forms that are a part of the project. </li>
<li><u>Module-Level Variable</u> - Variable that to all procedures on the form in which it is declared.</li>
<li><u>Local Variable</u> - Variable accessible only in the procedure which it was declared. </li>
</ul>
<br>
<h5>PROPERTIES</h5>
<br>
<p><u>Default Property</u> - Automatically selects a cmd button when the user presses the < ENTER > key. To make a cmd button a default button, you set its DEFAULT property to TRUE. Only one command button per form can have its default property set to true. When the program is run, this cmd button will be highlighted. </p>
<p><u>Cancel Property</u> - The button that is selected when the user presses the < ESC > key. To set a cmd button to the cancel button, you set its CANCEL property to TRUE. Only one cmd button per form can have its cancel button set to true.</p>
<p><u>TabStop Property</u> - Represents all controls on a form that can receive focus. If the TabStop property is TRUE, a control can receive focus. If it is FALSE, then it cannot.</p>
<br>
Some controls <b>can</b> receive focus, others cant. <br>
Txt Boxes, and cmd buttons <b>can</b> receive focus. <br>
Labels and images <b>cannot</b> receive focus.
<br>
<p><u>TabIndex Property</u></p> - Determines order of focus moves as the < TAB > key is pressed.</p>
<br>
<u>Name</u> - Used to assign the name of the control as it is known by the project. <BR>
<u>Caption</u> - The label that appears next to, in or on top of the control.<br>
<u>BackColor</u> - The background color of the control.<br>
<u>ForeColor</u> - The color of the text that appears on or next to the control.<br>
<u>Text</u> - The text which appears in a text box.<br>
<u>Alignment</u> - Determines justification of the text within a label or text box.<br>
<ul>
<li>0 - Left Justify</li>
<li>1 - Right Justify</li>
<li>2 - Center</li>
</ul>
<br>
<u>MultiLine</u> - Allows a string to be spread out among several lines instead of one line.<br>
<u>Font</u> - Allows you to set the font and font size of a control. <br>
<u>TabIndex</u> - Determines the order the focus moves as the < TAB > key is pressed. <br>
<u>Visible</u>- Property used to make a control visible or invisible<br><br>
<pre>
lblTotal.Visible = True 'label will be displayed on the form.
lblTotal.Visible = False 'label will not be displayed on the form.
</pre>
<br>
<u>FillStyle</u> - Primarily used with shapes, is used to fill the shape. Different options are available. <br>
<ul>
<li>0 - Solid</li>
<li>1 - Transparent </li>
<li>2 - Horizontal Line </li>
<li>3 - Vertical Line</li>
<li>4 - Upward Diagonal </li>
<li>5 - Downward Diagonal</li>
<li>6 - Cross</li>
<li>7 - Diagonal Cross</li>
</ul>
<br>
<u>EXAMPLES :</u>
<br>
<pre>
Square.FillStyle = 0
Square.FillStyle = 8
</pre>
<h4>Option Buttons... Formats</h4>
<br>
<p><u>Option Buttons</u> - Group of controls where only one can be selected at a time.</p>
<p><u>Check Boxes</u> - Group of controls where more than one can be selected at a time.</p>
<p><u>Frame</u> - Control that often acts as a container for a group of option buttons or check boxes.</p>
<p><u>Image</u> - Type of control that is used to hold a graphic.
Shape - Type of control used to place rectangles, squares, ovals and circles on a form.</p>
<br>
<h5>OPTION BUTTONS:</h5>
<br>
<p>The VALUE property of the option button is set TRUE if you want it to be selected, otherwise, FALSE. </p>
<br>
<p>Set the option buttons CAPTION property to the text you want to appear next to the option button.</p>
<br>
<p>When assigning an option button, a variable name always starts out with the prefix, opt</p>
<br>
<p>If you want an event to occur when you click on an opt button, you can double-click on the opt button to place code behind the opt button.</p>
<h5>CHECK BOXES:</h5>
<br>
More than one check box can be selected at a time.
<br>
<p>The VALUE property of the option button is set to CHECKED if you want it to be selected, otherwise, set it to UNCHECKED.</p>
<br>
A second option is to set the VALUE property on check boxes to a 0, 1 or 2.
<br>
<ul>
<li>0 - UnChecked</li>
<li>1 - Checked</li>
<li>2 - Grayed</li>
</ul>
<br>
Set the CAPTION property to the text you would like to appear next to the check box
<br>
When assigning a variable name, always start out with prefix chk
<br>
<p>If you want an event to occur when you click on a chk box, double-click the button to place code behind. </p>
<br>
<h5>IMAGES:</h5>
<p>Click on the images PICTURE property and locate the folder or drive where the picture is located.</p>
<br>
Select the pic you want and it will be placed on the form. <br>
<p>Set images STRECH property to TRUE. This will size the pic to the size of the control you have defined on the form.</p>
<br>
All controls that are images should begin with img prefix.
<br>
<h5>SHAPES:</h5>
<br>
The types of shapes and codes are shown below...
<br>
<ul>
<li>0 - Rectangle</li>
<li>1 - Square</li>
<li>2 - Oval</li>
<li>3 - Circle</li>
<li>4 - Rounded Rectangle</li>
<li>5 - Rounded Square</li>
</ul>
<br>
All controls that are shapes should begin with the shp prefix. <br>
<br>
<h5>LINES</h5>
<br>
<p>Use the crosshair pointer to drag a line across the screen. You may rotate the line in any direction and stretch it until releasing the move button.</p>
<br>
All line controls should begin with the lin prefix. <br>
<br>
<b>Determining Focus</b>
<br>
<p><u>Focus</u> - Refers to the currently selected control on the form. This can be indicated by an | - Beam, selected text, highlighted caption or dotted border. The control with focus is ready to receive input.</p>
<p><u>SetFocus</u> - A built-in function that when executed will move the cursor to the control and give that control focus. SetFocus can be used with txt boxes, cmd buttons, opt buttons and chk boxes.</p>
<br>
<i>Examples of setFocus:</i>
<br>
<pre>
txtNum1.setFocus 'will place cursor in the text box.
cmdCalc.setFocus 'will hilight this command button.
optBlue.setFocus 'will put dotted lines around this opt button.
</pre>
<br>
<h5>Working With Strings</h5>
<br>
<p><u>Concatenation</u> - Refers to combining two or more smaller strings into a larger string. In VB, you can either use the & or + symbols to do concatenation.</p>
<br>
<i>Examples of Concatenation:</i>
<br>
<pre>
StrFirstName + " " + strMiddleInitial + ". " + StrLastName
StrFirstName & " " + strMiddleInitial & ". " + StrLastName
</pre>
<br>
<h5>Formats</h5>
<br>
<b>CENTERING A FORM IN THE MIDDLE OF THE SCREEN:</b>
<br>
<p>(This code would go in the FORM LOAD [ by double clicking empty space on the form.])</p>
<br>
< Name of form > .Top = (Screen.Height - < Name of form > .Height)/2 <br>
< Name of Form > .Left = (Screen.Width - < Name of form > .Width)/2
<br><br>
<p><u> < Name of Form > </u> - This is the actual name of the form as you have saved it in your program.</p>
<p>If you named your form SQUARE, you would code the statement in the following way: </p>
<br>
<pre>
SQUARE.Top = (Screen.Height - SQUARE.Height)/2
SQUARE.Left = (Screen.Height - SQUARE.Width)/2
</pre>
<h4>Input Boxes, Message and List Boxes</h4>
<br>
<br>
<p><u>Input Box</u> - A function that will display a message and allow the user to enter information in a text box. In the Input box you can display a message called a prompt, which will help the user decide what information he needs to enter in the text box.</p>
<br>
<p>The input box will have a txt box with the prompt above it and two command buttons, OK and CANCEL. The OK will accept whatever input the user enters and place it in the variable on the left hand side of the equals sign. The CANCLEL button will ignore any input entered by the user and return the user back to the form that is currently open.</p>
<br>
Input Boxes are often used when one wants to retrieve records from files.<br>
<br>
VariableName = InputBox("Prompt","Title")<br>
<br>
<u>Examples:</u><br>
<pre>
StrName = InputBox("Enter your name", "Sponge Bob")
StrCareer = InputBox("Enter your workplace", "Nickelodeon")
</pre>
<br>
<p>The prompt must be enclosed in quotes. The title must also be in quotes and will appear in the Title Bar of the Input Box. If no title is entered, the title of the project will appear in the Title Bar.</p>
<br>
Input Boxes can appear in the following places: <br>
<br>
<ul>
<li>Form_Load</li>
<li>Command Buttons</li>
<li>Option Buttons</li>
<li>Check Boxes</li>
</ul>
<br>
<h5>Form Load</h5>
<br>
<p><u>Form_Load</u> - Code executed as the project is loading. The first time a form is displayed in a project, VB generates an event knows as FORM_LOAD. Any code in Form Load is then executed.</p>
<br>
Things that are done in the FORM_LOAD section of a VB program include the following:<br>
<ul>
<li>Code to center the form in the middle of the screen</li>
<li>Code to initialize variables</li>
<li>Code for input boxes so the user can enter info.</li>
<li>Display info in labels on the form.</li>
</ul>
<br>
<p>To get the FORM_LOAD event to enter code, again; double click on an empty area on the form. The FORM_LOAD event begins and ends with the following procedure headings:</p>
<br>
Private Sub Form_Load()<br>
<br>
< Place Code Here > <br>
End Sub <br>
<br>
<br>
<i>Example of code in the FORM_LOAD event:</i><br>
<pre>
Private Sub Form_Load()
InputBoxes.Top = Screen.Height - InputBoxes.Height)/2
InputBoxes.Left = Screen.Width - InputBoxes.Width)/2
StrName = InputBox("Enter your name", " Nickelodeon Inc.")
LblName = StrName
End Sub
</pre>
<br>
<h5>Message Boxes and List Boxes</h5>
<br>
<p><u>Message Box</u> - A special type of VB statement/function that displays a window in which you can display a message to a user. The message box can be a statement or function and has the name, MsgBox.</p>
<br>
The following can be displayed to the user with a Message Box:<br>
<ul>
<li>Message</li>
<li>Optional Icon</li>
<li>Title Bar Caption</li>
<li>Command Buttons</li>
</ul>
<br>
<h5>MESSAGE BOX STATEMENT</h5>
<br>
<p>The message box statement is designed to be on a line by itself. The syntax of the MsgBox is show here:</p>
<br>
MsgBox < "Prompt" > , < Buttons/Icons > , < "Caption" > <br>
<pre>
MsgBox "Sponge Bob is -Jester’s brother!",vbInformation, "Nickelodeon Inc."
</pre>
<br>
<p><u>Prompt</u> - The message you want to appear in the message box.
Buttons/Icons - This determines what command buttons and/or icons that will appear on the message box. (This portion is optional.)</p>
<u>Caption</u> - This is the caption that will appear on the title bar of the message.<br>
<br>
<h5>MESSAGE BOX FUNCTION:</h5>
<br>
<p>The message box function will appear on the right hand side of the equals sign. Also, if your msg box is a function, you must enclose arguments in ( ).</p>
<br>
VarName = MsgBox( < "Prompt" > , < Buttons/Icons > , < "Caption" > ) <br>
<pre>
IntRes = MsgBox("Are my fingers tired?", vbYesNo + vbQuestion,"My Question")
</pre>
<br>
<i>Sample code using Message Boxes:</i> <br>
<pre>
Private Sub cmdCheck_Click()
If val(txtNum1) > val(txtNum2) Then
MsgBox "First Number is greater than second", vbInformation, "Comparing Numbers"
ElseIf val(txtNum2) > val(txtNum1) Then
MsgBox "Second Number is greater than the first", vbInformation, "Comparing Numbers"
Else
MsgBox "The two numbers are equal", vbInformation, "Comparing Numbers"
End If
End Sub
</pre>
<br>
<h5>LIST BOXES:</h5>
<br>
<p><u>List Box</u> - A type of control used to hold a list of items from which the user can select one item from the list. You should use the lst prefix when naming list boxes. </p>
<br>
Items can be added to a list box in two ways.<br>
<br>
Using the LIST property for the list control <br>
Using the AddItem method<br>
<br>
< object_name > .AddItem < Value > <br>
<br>
<p>If your values are strings, they must be enclosed in double quotes. When items are added to the list, they are given an index. The first item added to the list will have an index of zero the 2nd of one, and so forth.</p>
<br>
<i>Example of - ADDING ITEMS TO THE LIST:</i><br>
<pre>
lstSchools.AddItem "Harvard"
lstSchools.Additem "Yale"
lstSchools.Additem "Princeton"
lstSchools.Additem "Brown"
lstSchools.Additem "Cornell"
lstSchools.ListIndex = 3
</pre>
<br>
This will highlight the item "Brown". Harvard will have an index of 0, Yale of 1 and so forth. <br>
<br>
<h5>ListIndex PROPERTY :</h5>
<br>
<p>The listIndex property will highlight an item in the list when the program is run. Only one item can be set with the listIndex property. The format of List Index property is shown here:</p>
<br>
< control > .ListIndex = 3 'This will highlight the 4th item in the list. <br>
<pre>
lstSchools.ListIndex = 3 'This will highlight the 4th school in lstSchools.
</pre>
<br>
<h5>Sorted PROPERTY:</h5>
<br>
<p>This will sort the items in your list alphabetically if the SORTED property is set to TRUE. This will also re-index the items in your list. If the ListIndex property is used, the highlighted item will change.</p>
<br>
<h5>Clear METHOD:</h5>
<br>
This will empty out the contents of a list box...<br>
<br>
< control > .Clear <br>
<pre>
lstSchools.Clear
</pre>
<br>
<h5>RemoveItem METHOD:</h5>
<br>
<p>This will remove one item from the list. When using RemoveItem, youmust include the index of the item to remove. </p>
<br>
< control > .RemoveItem < index > <br>
<pre>
LstSchools.RemoveItem 3
</pre>
<p>This will remove the item in the lstSchools list box with the index of 3, which is the 4th item in the list.</p>
<br>
<p><u>LstCount</u> - The listCount property is used to hold the number of items in the list. ListCount will always be one more than the highest element in the list.</p>
<h4>Sub Procedures, Random Numbers</h4><br>
<br>
<p><u>Procedure</u> - A unit of code that performs a specific task and can be called from other locations of the program.</p>
<br>
<h5>PURPOSES OF PROCEDURES:</h5><br>
<br>
<ul>
<li>Breaks large sections of code into smaller units of code that perform a specific task.</li>
<li>Makes it easier to debug and maintain program.</li>
<li>Cuts down on the amount of code that has to be written and eliminates duplication of code.</li>
</ul>
<br>
<h5>TYPES OF PROCEDURES:</h5>
<br>
<p><u>Sub Procedure</u> - A procedure that performs a task but DOES NOT return values back to the calling module.</p>
<p><u>Function Procedure</u> - A procedure that performs a task and RETURNS a value back to the calling module. With function procedures, the value is returned back to the calling module using the function name.</p>
<br>
<h5>CREATING A NEW SUB PROCEDURE:</h5><br>
<br>
<ol>
<li>Display the code window for the form</li>
<li>Select ADD PROCEDURE from the TOOLS menu</li>
<li>Enter the name of the procedure in the text box next to where it says NAME</li>
<li>Select PRIVATE for SCOPE</li>
<li>Click OK</li>
<li>You will be given the procedure shell. Type in the contents of your procedure</li>
<li>Click on the code button to exit the procedure</li>
</ol>
<br>
<i>Example of a Sub Procedure:</i><br>
<pre>
Private Sub AddNumbers()
Sum = val(txtNum1) + val(txtNum2)
LblSum = Sum
End Sub
</pre><br>
<i>Example of a Sub Procedure Call: </i><br>
<pre>
Private Sub cmdCalculate_Click()
AddNumbers
End Sub
</pre><br>
<i>Example of a Function Procedure:</i><br>
<pre>
Private Function AddNumbers()
AddNumbers = val(txtNum1) + val(txtNum2)
End Function
</pre><br>
<i>Example of a Function Procedure Call:</i><br>
<pre>
Private Sub cmdCalculate_Click()
LblSum = AddNumbers()
End Sub
</pre>
<br>
<h5>WITH STATEMENT</h5><br>
<br>
<p>The WITH keyword allows you to cut down on coding when dealing with properties of controls.</p>
<br>
<i>Syntax of WITH statement:</i><br>
<br>
With < control > <br>
. < property1 > = < value1 > <br>
. < property2 > = < value2 > <br>
. < property3 > = < value3 > <br>
End With <br>
<br>
<i>Assigning controls without the WITH KEYWORD:</i>
<br>
<pre>
lblEmployee.Font.Name = dlgCommon.FontName
lblEmployee.Font.Bold = dlgCommon.FontBold
lblEmployee.Font.Italic = dlgCommon.FontItalic
</pre>
<br>
<i>Assigning controls using the WITH KEYWORD:</i>
<br>
<pre>
With lblEmployee.Font
.Name = dlgCommon.FontName
.Bold = dlgCommon.FontBold
.Italic = dlgCommon.FontItalic
End With
</pre>
<br>
<h5>COMMON DIALOG CONTROL</h5>
<br>
<p>Allows your project to use the dialog boxes that are provided as part of the Windows environment to set properties for a control such as font, font size and color.</p>
<br>
<h5>FEATURES OF THE DIALOG CONTROL:</h5>
<br>
<ul>
<li>You only need common dialog control on your form</li>
<li>You cannot change the controls size </li>
<li>The location of the control does NOT matter </li>
<li>The control will be invisible when the program runs </li>
<li>Dialog controls are stored with an extension of .ocx </li>
<li>When naming Dialog controls, begin with a prefix of dlg </li>
<li>The common dialog box may not appear in your toolbox. It is a custom * control and will need to be added to your project before you can use it.</li>
</ul>
<br>
<h5>RETRIEVING THE COMMON DIALOG CONTROL:</h5>
<br>
<ol>
<li>click on Project</li>
<li>click on Components </li>
<li>Scroll down and find Microsoft Common Dialog Control 6.0 </li>
<li>Click the check box next to it to select it. </li>
<li>Click OK and it will be placed in your tool box. </li>
<li>Click on Dialog Control and place it on your form. </li>
</ol>
<br>
<h5>CHANGING FONTS:</h5>
<br>
<i>Example SYNTAX of SETTING FONTS:</i>
<br><br>
With dlgCommon (assuming that you named it that)<br>
.Flags = cdlCFScreenFonts ‘Loads different fonts into memory<br>
.ShowFont<br>
End With<br>
<pre>
With txtName.Font
.Bold = dlgCommon.FontBold
.Italic = dlgCommon.FontItalic
.Name = dlgCommon.FontName
.Size = dlgCommon.FontSize
End With
</pre>
<br>
<h5>CHANGING COLOR:</h5><br>
<br>
<i>Example SYNTAX of CHANGING COLOR:</i>
<br>
<pre>
dlgCommon.ShowColor 'Brings up the color box
txtName.ForeColor = dlgCommon.Color 'Applies it to the font
</pre>
<br>
<h5>RANDOM NUMBERS</h5>
<br><br>
<u>Rnd</u> - Generates a random number between 0 and 1.<br>
<p><u>Randomize</u> - Tells VB to randomly generate numbers for the rnd statement. Using randomize will allow the rnd statement to generate an entirely random set of numbers that do not follow any recognizable pattern.</p>
<br>
<i>Generating Numbers between 1 and 10:</i>
<br>
<pre>
Num = Int((10 - 1 + 1)* rnd + 1)
</pre>
<br>
<i>Generating Numbers between 1 and 100:</i>
<br>
<pre>
Num = Int((100 - 10 + 1)* rnd + 1)
</pre>
<br>
<h5>INT() FUNCTION:</h5>
<br>
<p>Converts a floating point value to an integer by truncating off any remainder that the number has. ( INT(4.656) will return a value of 4. )</p>
<h4>Menus, Combo Boxes and QB Color</h4>
<br>
<p><u>Menu</u> - A drop-down list of items displayed below a menu name on the screen from which you selected one item.</p>
<br>
<p>In Windows and VB a menu consists of a menu bar with menu names, each of which drops down to display a list of menu commands. You can use menu commands in place of or in addition to the command buttons to activate a procedure.</p>
<br>
<p>Menu commands are actually controls and have events and properties. Each menu command has a Name property and a Click event, similar to a command button. To create a menu for your form, you will use the Visual Basic Menu Editor, which accessible by pressing < CTRL > + E or Click on Menu Editor Icon.</p>
<br>
<h5>PARTS OF THE MENU:</h5>
<br>
<u>Caption</u> - Holds the words you want to appear on the screen.<br>
<p><u>Name</u> - Indicates the name of the menu control and what the control is referred to by the program. When naming, start out with the prefix, "mnu" . For example, if you had a menu control for FILE, you would name it, mnuFile</p>
<p><u>SUBMENU</u> - A list of commands that appear underneath a menu command, a menu within a menu. To create a submenu, press the right arrow key to move to the next level.</p>
<p><u>Menu List Box</u> - Shows the list of all menu items that have been created and the indication levels. You can move up, down, left and right by clicking on the name of the menu item and then clicking on one of the four arrow buttons.</p>
<p><u>Separator Bars</u> - A horizontal line that separates one menu from another. To define a separator bar, type a single hyphen ( - ) for the caption and give it a name. Even though you can never reference the separator bar in the code, you still have to give it a name. If you have more than one separator bar, each one has to have a unique name. ( mnuSep1,mnuSep2,...ect. )</p>
<p><u>Short Cut Keys</u> - You can create a keyboard short cut key for a menu item when its created. Select a short cut key for your menu item by selecting it from the list provided. The purpose of the short cut is to give you an alternative to going through a menu to perform a task.</p>
<p><u>Checked</u> - If you want a check mark to appear next to an item in your menu, you would put a check here. Check marks are normally for options that you want to be toggled on and off.</p>
<p><u>Enabled</u> - If you want the user to be able to select a menu item, this will be checked. By default, all are enabled. If unchecked, this menu item will be dimmed out. You can also change this property in the actual code. ( mnuStar.Enabled = False )</p>
<p><u>Visible</u> - Determines whether a menu item is displayed on the screen or not. If there is a check in this check box, the menu item will be displayed on the screen. If the box is not checked, the menu item will not be displayed. This property can also be changed in the actual code. ( mnuStar.Visible = True )</p>
<p><u>Insert</u> - Click on this if you want to insert an item into the middle of a menu. The item will be inserted at the location of the current item. All items below that point will be moved down.</p>
<p><u>Delete</u> - Will delete the highlighted menu, Cancel will close the menu editor without saving changes and OK will save changes to the menu and close the menu editor.</p>
<br>
<h5>COMBO BOXES</h5>
<br>
<p><u>Combo Box</u> - A type of control used to hold a list of items from which the user selects one from a pull down menu.</p>
<br><br>
<h5>DIFFERENCES BETWEEN A LIST BOX AND A COMBO BOX:</h5>
<br>
<ul>
<li>With a list box, there is no pull down menu, user scrolls through the list using scroll bars and selects the item of his/her/OtHeR choice. With a combo box, the user click on a down arror, a pull down menu is displayed and the user chooses by clicking one on the list.</li>
<li>You can type your own entry into the list with a combo box. With a list box, you cannot type in your own entries.</li>
</ul>
<br>
Like list boxes, items can be added to a combo box in two ways...<br>
<br>
<ul>
<li>Using the List property</li>
<li>Using the AddItem method</li>
</ul>
<br>
<h5>ADDING ITEMS TO THE LIST USING THE LIST PROPERTY:</h5>
<br>
<ol>
<li>Scroll through the properties window to the LIST property </li>
<li>Click on the down arror to drop down an empty list </li>
<li>Type your first item and press < CTRL > < ENTER > </li>
<li>Continue entering items as indicated in 3 until finished </li>
<li>Press < ENTER > or click outside of list box to complete operation.</li>
</ol
<br>
<h5>ADDING ITEMS TO THE LIST USING THE ADDITEM PROPERTY:</h5>
<br>
<ol>
<li>Double click on the form to open FORM_LOAD </li>
<li>Enter your items using the below format :</li>
</ol>
<br>
< object_name > .AddItem " < Value > " <br>
<br>
<h5>ADDING ITEMS INTO THE ItemData ARRAY:</h5><br>
<br>
< object_name > .AddItem " < Value > " <br>
< object_name > .ItemData( < object_name > .NewIndex) = < Value > <br>
<br>
<h5>ADDING ITEMS TO THE LIST:</h5><br>
<br>
Puttings items into the list of a Combo Box...<br>
<pre>
cboSchool.AddItem "Harvard"
cboSchool.AddItem "Yale"
cboSchool.AddItem "Princeton"
cboSchool.AddItem "Brown"
cboSchool.AddItem "Cornell"
</pre>
<br>
Harvard will have the index of 0, Yale of 1 and so on<br>
<br>
<i>Entering Items into the ItemData array of a Combo Box...</i><br>
<pre>
cboOffice.AddItem "Paper"
cboOffice.ItemData(cboOffice.NewIndex) = 300
cboOffice.AddItem "Cartrige"
cboOffice.ItemData(cboOffice.NewIndex) = 3200
cboOffice.AddItem "Folders"
cboOffice.ItemData(cboOffice.NewIndex) = 250
cboOffice.AddItem "Binder"
cboOffice.ItemData(cboOffice.NewIndex) = 400
</pre>
<br>
<h5>TEXT PROPERTY:</h5><br>
<br>
The Text property refers to the actual item that is currently selected in the list.<br>
<br>
< label > = < control > .Text <br>
<pre>
lblFood = cboFood.Text
</pre>
<br>
<p>ListIndex works as the same as List Boxes. Same as NewIndex, ItemData, Sorted, Clear, RemoveItem, ListCout...ect.</p>
<br>
<p>When naming Combo Boxes, use the prefix, "cbo". For example, if you wanted a combo box named School, you would name it, cboSchool.</p>
<br>
<h5>QB COLOR</h5>
<br>
<p><u>QBCOLOR</u> - is a built-in function in VB that will display 15 different colors. The QBCOLOR function has one argument, which is the index of the color to be displayed. The index can range from values 0 to 15. What the heck, here are the values for you:</font></p>
<br>
<table width="33%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolor="#FFFFFF">
 <tr bordercolor="#FFFFFF">
 <td bgcolor="#CCCCCC" valign="middle" align="center" font color="#000000">-Index-</font></td>
 <td bgcolor="#999999" valign="middle" align="center">
 <div align="center"><font color="#000000">-Color-</font></div>
 </td>
 </tr>
 <tr>
 <td align="left" valign="top" bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>0</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Black</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>1</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Blue</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>2</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Green</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>3</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Cyan</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>4</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Red</div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>5</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Magenta</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>6</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Yellow</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>7</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">White</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>8</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Gray</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>9</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Light Blue</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>10</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Light Green</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>11</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Light Cyan</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>12</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Light Red</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>13</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Light Magenta</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>14</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Light Yellow</font></div>
 </td>
 </tr>
 <tr>
 <td bgcolor="#E8F8FF">
 <div align="left"><font color="#006699"><b>15</b></font></div>
 </td>
 <td bgcolor="#006699" bordercolor="#006699">
 <div align="left"><font color="#FFFFFF">Bright White</font></div>
 </td>
 </tr>
 </table>
<br>
<font size="2">
<i>Example:</i><br>
<pre>
ShpRectangle.FillColor = QBColor(1) ‘Will be shown in blue
ShpCircle.FillColor = QBColor(6) ‘ Will be shown in yellow
</pre>
<h4>Control Arrays, Scroll Bars and RGB Function</h4>
<br>
<br>
<u>Array</u> - A group of data items that are referred to by the same name. <br>
<u>Control Array</u> - A group of controls sharing the same name and event procedures. <br>
<u>Index</u> - A variable used to refer to each element in the control array.<br>
<br>
<i>Control Arrays can be used with the following controls:</i> <br>
<br>
<ul>
<li>Text Boxes </li>
<li>Option Buttons </li>
<li>Check Boxes </li>
<li>Labels </li>
<li>Command Buttons </li>
</ul>
<br>
Control Arrays are most commonly used with text boxes and option buttons.<br>
<br>
<h5>CREATING A CONTROL ARRAY:</h5>
<br>
<ol>
<li>drop a control on the form and assign it a name. </li>
<li>drop a second control on the form and assign it the SAME name. </li>
<li>You will be given a message saying that you already have a control named ‘whatever’ and asks you if you want to create a control array.</li>
<li>Press "Yes" and this will create your control array.</li>
<li>Now, for each subsequent item added, give it the same name also.</li>
<li>Once you create the control array, each item within the array will have an index. The first will be an index of zero, 2nd of 1, 3rd of 2 and so forth...</li>
</ol>
<br>
<br>
<i>Example Control Array using a Select Case:</i><br>
<p>This example is if you had 6 option buttons with names of colors and a Rounded Square shape that will fill the color of the selected option button. </p>
<br>
<pre>
Private Sub optColors_Click(Index As Integer)
Select Case Index
Case 0: shpRoundSquare.FillColor = vbBlack
Case 1: shpRoundSquare.FillColor = vbBlue
Case 2: shpRoundSquare.FillColor = vbYellow
Case 3: shpRoundSquare.FillColor = vbGreen
Case 4: shpRoundSquare.FillColor = vbRed
Case 5: shpRoundSquare.FillColor = vbMagenta
End Select
End Sub
</pre>
<br>
<h5>SUM AND AVERAGE OF TEST SCORES USING CONTROL ARRAYS:</h5>
<br>
<br>
<p>Here is another example of a control array. This will be if you wanted to figure the sum and average of test scores.</p>
<br>
<pre>
'Calculating Sum and Average
Private Sub cmdCalculate_Click()
For X = 0 to 9
Sum = Sum + val(txtTest(X))
Next X
Average = Sum / 10
End Sub
'Clearing text boxes and totals
Private Sub cmdClear_Click()
For Z = 0 to 9
TxtTest(Z) = ""
Next Z
LblSum = ""
LblAvg = ""
Sum = 0
Average = 0
End Sub
</pre>
<br>
<h5>Scroll Bars and RGB Function</h5><br>
<br>
<p><u>Scroll Bar</u> - A control that allows you to see hidden information on the screen. This information can be text, icons or controls.</p>
<p>In VB a scroll bar is used to represent a range of values. Scroll bars are also used to control sound level, color, size and other values that can be changed in small amounts or large increments.</p>
<p><u>Scroll Box</u> - The little square, which appears inside the scroll bar. Pressing down on the scroll box changes the value property of the scroll bar. Another name for the scroll box is the thumb. </p>
<br>
<h5>PROPERTIES OF THE SCROLL BAR:</h5>
<br>
<ul>
<li><u>Min</u> - The smallest value the scroll bar can take on <li>
<li><u>Max</u> - The largest value the scroll bar can take on <li>
<li><u>Small Change</u> - The distance to move when the user clicks on the scroll arrows <li>
<li><u>Large Change</u> - The distance to move when the user clicks on the gray area of the scroll bar <li>
<li><u>Value</u> - Indicates the current position of the scroll bar and its corresponding value within the scroll bar.<li>
</ul>
<br>
<h5>TYPES OF SCROLL BARS:</h5><br>
<br>
<ul>
<li>Vertical Scroll Bar - Begins with the prefix "vsb" </li>
<li>Horizontal Scroll Bar - Begins with a prefix "hsb" </li>
</ul>
<br>
<h5>EVENTS OF THE SCROLL BAR:</h5> <br>
<br>
<u>Change Event</u> - Occurs when the user clicks on the gray area of the scroll bar. <br>
<u>Scroll Event</u> - Occurs when the user drags the scroll box.<br>
<br>
<p>As soon as the user releases the mouse botton, the scroll event ceases and a change event occurs. When you write code for the scroll bar, you will want to write code for both Change even and the Scroll event. </p>
<br>
<h5>SAMPLE PROGRAM USING HORIZONTAL SCROLL BAR:</h5><br>
<pre>
Private Sub cmdExit_Click()
End
End Sub
Private Sub Form_Load()
Hscroll.Top = (Screen.Height - Hscroll.Height)/2
Hscroll.Left = (Screen.Width - Hscroll.Width)/2
End Sub
Private Sub hsbScroll_Change()
LblValue = hsbScroll.Value
End Sub
Private Sub hsbScroll_Scroll()
LblValue = hsbScroll.Value
End Sub
</pre>
<br>
<h5>SAMPLE PROGRAM USING VERTICAL SCROLL BAR:</h5>
<br>
<pre>
Private Sub cmdExit_Click()
End
End Sub
Private Sub Form_Load()
Vscroll.Top = (Screen.Height - Vscroll.Height)/2
Vscroll.Left = (Screen.Width - Vscroll.Width)/2
End Sub
Private Sub vsbScroll_Change()
LblValue = vsbScroll.Value
End Sub
Private Sub vsbScroll_Scroll()
LblValue = vsbScroll.Value
End Sub
</pre>
<br>
<h5>RGB FUNCTION</h5><br>
<br>
<p>The RGB function specifies the quantities of red, green and blue for a large variety of colors. The value for each color ranges from 0 to 255 with 0 being the least intense and 255 being the most intense. The color arguments are in the same order as their letters in the function name, red first, then green and then finally blue. You can use the RGB function to assign the color to a property or specify the color in a graphics method. </p>
<br>
Chosen_Color = RGB(RedValue, GreenValue,BlueValue)<br>
<br>
<i>Examples:</i><br>
<br>
<pre>
RGB(0,0,0) 'would be BLACK
RGB(255,255,255) 'would be BRIGHT WHITE
RGB(255,0,0) 'would be RED
RGB(0,255,0) 'would be GREEN
RGB(0,0,255) 'would be BLUE
</pre>
<h4>Multiple Forms, ME Keyword and Conclusion</h4>
<br>
<br>
<p>A VB project can consist of several forms. Each form has its own window, form load section, general declarations section and code window. All of the forms in a project are tied together under the project name. When you run the program, you can switch back and forth between different forms.</p>
<br>
<u>StartUp Form</u> - The first form a project displays when the project is loaded.<br>
<u>Load</u> - Loads a form into the computers memory (does not display on screen.)<br>
<br>
 · Load < form_name > <br>
<br>
<p><u>Unload</u> - Unloads the form from the computers memory. Also, if you’re running an execss amount of forms, you should unload them to save up on memory.</p>
<br>
· Unload < form_name > <br>
<br>
<u>Show</u> - Display the loaded form on the screen. <br>
<br>
· < form_name > .Show <br>
<br>
<p><u>Hide</u> - Makes the form disappear. The form will still be loaded into memory. It just will not be displayed on the screen.</p>
<br>
· *lt; form_name > .Hide <br>
<br>
<h5>ME KEYWORD:</h5>
<br>
<p>You can refer to the current form by using the special keyword ME. Me acts like a variable and refers to the form that is currently active. You can use Me in place of the form name when coding statements and methods.</p>
<br>
<i>Examples:</i>
<br>
<pre>
Unload Me 'Unloads the current form that is executing code
Me.Hide 'Hides the form currently executing code
Me.Show 'Shows the form currently executing code
</pre>
<br>
<h5>Conclusion:</h5>
<br>
<p>Alright! Wow, I believe that is an acceptionally well start on Visual Basic 6.0. This tutorial does not cover many things, however you can learn a lot just by playing around with the program. I am sure that I left a lot of things out that I should have put in here, and discussed things I could have left out- but this is my first tutorial, so take it easy! Anyways, hope you enjoy and are ready to code and design in Visual Basic!!!</p>
<br>
<b>-bs0d | www.allsyntax.com</b><br></font>

