# VBA Introduction
This is an introductory guide to Visual Basics for Applications (VBA). At its core, VBA is a scripting languauge that provides users with the ability to control the Microsoft Office envirnment programmatically.

Like most programming languages, VBA has a number of programming constructs that help to to extend its power and flexibility. Additionally it has a gentle learning curve that makes it a suitable first language for beginners in the world of programming.

This guide will offer a gentle introduction to VBA and its constructs. It is intended to provide readers with a solid foundation upon which to explore and extend their skills.

### Invoking the VBA Editor
The Microsoft Office suite provides a built-in code editor that allows users to code instructions relevant to the application. There are several ways to invoke your VBA Editor, depending on your operating system.<br>
For Windows
- Option 1: While holding the **Alt** key, press the function **F11** key (**Alt + F11**)
- Option 2: Select the Developer tab, then press the Visual Basic button to activate the VBA Editor

For MAC
- Option 1: While holding the **fn** + **Alt** keys, press the function **F11** key (**Alt + F11**)
- Option 2: Select the Developer tab, then press the Visual Basic button to activate the VBA Editor

hint: The Developer Menu is not viosible by default. if you need to activate it, do the following:
  From the menu choices choose *File* - *Options* - *Customize Ribbon* , then check Developer Option from the column on right. 

Below is a sample of the VBA Editor window that appears when we use one of the above Options to launch the editor.
![VBA IDE](https://github.com/informidas/vba-basic-documentation/blob/master/VBA_IDE.PNG "sample VBA Editor screen")


---

Here a few key considerations before we begin our sample coding:
- To begin coding, we will enter our instructions in the blank area below general.
- These instructions that we enter are called statements or commands.
- Each new statement / command is placed on a separate line
- to end and instruction we simply press the Enter / retun key
  
### VBA Constructs
VBA provides some the most useful programming constructs, many of which can be found in other popular programming language. 
##### Declaring a Variable
In order to use a variable in VBA we define it as follows:
*Dim variableName as variableType*
Some of the most frequently used Variable Types are:

* string
* integer
* long
* double
* single
* boolean
* array
* date
* decimal
* byte
* currency

##### Generating Comments
It is a good coding practice to include commenst in your code. Comments provide a way for others reviewing your code to understand the intent of the statements in particular and your program in general


#### Subroutines
A Subroutine is a block of code (i.e. series of vba statements or commands). This subroutine when executed will run all statements in the block.

Creating a subroutine begins with the keyword *Sub* and ends with the keywords *End Sub* . Below is an example of a subroutine.

>
> ##### Declaring a Subroutine
> Sub HelloWorld()
>
> End Sub
