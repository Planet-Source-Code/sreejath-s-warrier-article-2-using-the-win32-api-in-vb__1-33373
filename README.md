<div align="center">

## Article \#2 Using the Win32 API in VB


</div>

### Description

We'll see how to access the Win32 API from Visual Basic. We'll start with the basics of declaring and using an API function. We'll discuss parameter passing, passing by value, passing by reference, aliasing etc and Develop a small Utility.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sreejath S\. Warrier](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sreejath-s-warrier.md)
**Level**          |Intermediate
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sreejath-s-warrier-article-2-using-the-win32-api-in-vb__1-33373/archive/master.zip)





### Source Code

<B>Article #2 Using the Win32 API in VB</B><P>
<B>Foreword</B><P>
In this article we'll see how to access the Win32 API from the Visual Basic environment. We'll start with the basics of declaring and using an API function. From there we'll move on to discussing parameter passing, passing by value, passing by reference, aliasing etc. in the course of this we develop a small utility that displays the computer name in a message box.
<P>
As discussed in Article#1, the API functions are declared in DLLs located in the Windows System directory. You can type in the declaration of an API just as you do with any other function exported from a DLL, but VB has provided an easier way to do it. It is called API Text Viewer. This was available as a standalone application in VB5. In VB6 this can be run as a standalone app or as an addin.
<P>
To have some API function(s) declared in your project, just launch API Text Viewer, open Win32Api.txt (or .MDB if you have converted it into a database to speed it up), choose Declares, find the function, and click Add and then Copy. Go to your project and paste it in. Do the same to have a predefined constant or type. This saves quite lot of typing. Also it eliminates the errors caused due to mistyping.
<P>
However, you need to watch out for these common glitches:
<P>
•	If you copy the declarations right from the API Text Viewer, the scope of the function is not specified. Since the scope is Public by default, the function has public scope. But VB does not allow public Declare statements within form modules. So if we are copying the Declare statements into a form module, they should be explicitly specified as Private.
<P>
This strategy has its own problems as this makes the declarations invisible outside the module in which they are defined. So the functions should be declared in every module that needs to use them. Declaring them in a separate .BAS module where they can be declared as Public and are therefore accessible to all the modules eliminates this duplication. This also permits easier code maintenance.
<P>
•	Another common error is the Ambiguous Name detected error. Put simply this means that the module contains more than one function or constant having the same name, causing the ambiguity. Since most of the functions are Aliased (see Aliasing, later in this article), they are referred to by different names. In such cases we can change the function name at the declaration and invocation points and the code will work fine.
<P>
•	There is one more error that you need to keep a really sharp eye out for, mainly because it comes from an unexpected source. The Win32api.txt file supplied by Microsoft contains some errors, which can cause you endless hours of frustration. Since we copy it from the API Text Viewer, the Declaration would be the last place we'd expect an error to arise especially as the file itself is a Microsoft product. But sadly this file contains a number of errors that are particularly hard to track down. There's nothing much you can do about it, except of course to correct any that come
to your attention and consult a good API guide or SDK documentation when things don't seem to go as expected.
<P>
Now we'll see a simple declaration and invocation example. For this purpose let us build an application that finds out your computer name, i.e. the name you have given your computer.
<P>
Ready?
<P>
OK, here we go!
<P>
First, load VB and create a new standard EXE project. A form should have been added to the project by default. In the General | Declarations section of the form add the following code:
<P><code>
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA"
(ByVal sBuffer As String, lSize As Long) As Long
<P>
'Now add the following code to an appropriate procedure. <P>Here I've used the
'Form_Load event<P>
Private Sub Form_Load ()<P>
Dim strString As String<P>
'Create a buffer<P>
strString = String(255, Chr$(0))<P>
'Get the computer name<P>
GetComputerName strString, 255<P>
'Remove the unnecessary Chr$(0)<P>
strString = Left$(strString, InStr(1, strString, Chr$(0)) - 1)<P>
'Show the computer name<P>
MsgBox strString<P>
End Sub<P>
</code>
<P>
Please examine the Declaration part carefully. <P>Straightaway, we can notice several things.
<P>First of all we see that the scope of the Declaration is private. This is because VB does not allow Public Declare statements in object modules. <P>Next we see that this is the declaration for a Function and that the name by which the function is referred in this program is GetComputerName. This is NOT the real name of the function as will become clear when we discuss aliasing in a short while. The next two words are interesting.<P>
Lib "kernel32"<P>
The Lib Keyword is used to specify the path and file name of the Library(DLL) which encapsulates the function.<P>
Lib "kernel32" means that the specified API function is found in the kernel32.dll file.
Strictly speaking, it should be "kernel32.dll" but the extension name can be skipped. Also, the path can be skipped if the file is located in the System folder.<P>
<P>
<B>Aliasing</B><P>
Aliasing is a technique used to call a function with a different name than its original name as defined in the DLL to avoid naming problems ambiguities etc. Occasionally, a DLL procedure has a name that is not a legal identifier. It might have an invalid character (such as a hyphen), or the name might be the same as a Visual Basic keyword (such as GetObject). When this is the case, use the Alias keyword to specify the illegal procedure name. I.e. For example, some procedures in the operating environment DLLs begin with an underscore character. While you can use an underscore in a Visual Basic identifier, you cannot begin an identifier with an underscore. To use one of these procedures, you first declare the function with a legal name, then use the Alias clause to reference the procedure's real name.
<P>
We can also use the Alias clause to change a procedure name whenever it's convenient say, to remove the ambiguity problem discussed earlier in this article. If you do substitute your own names for procedures, make sure that you thoroughly document the changes so that your code can be maintained at a later date.
<P>
In our example therefore, <B><code>"GetComputerNameA"</code></B> is the name of the function and
GetComputerName is its alias. This could be confusing the first time around. Don't worry; it's simple once you get the hang of it.
<P><B>Arguments</B><P>
This function takes two arguments. The first argument is a string and is passed ByVal. The second argument is a long value and is passed ByRef. Wait a minute! What are these ByVal and ByRef arguments and what is their significance? We'll see that in a minute. Now let us see the argument list:
<P>
lpBuffer - Points to a buffer to receive the null-terminated character string containing the computer name.<P>
nSize - Points to a variable that specifies the maximum size, in characters, of the buffer. This value
should be large enough to contain MAX_COMPUTERNAME_LENGTH + 1 characters.
<P>
<B>ByVal and ByRef</B><P>
If you do API programming for any length of time, it is impossible to avoid these two terms. They refer to two methods used to pass arguments to functions.<P>
<B>ByVal</B> means that the argument is passed by value. I.e. a copy of the argument is made and this copy is passed to the function. Any changes that the function makes to its own copy of the data are not reflected to the original value. <P>
<B>ByRef</B> means that the address of the variable is passed as argument to the function. So any changes the function makes to the argument is immediately reflected in the original variable. In passing arguments to API functions, ByRef is the default, whereas if you need to pass an argument ByVal, it needs to be explicitly specified.<P>
Now let us analyse the rest of the code. In the Form_Load event procedure we first create a string buffer, strString. Then we pre-size (or stuff) it with enough Chr$(0) (space) so that it is larger than the maximum length that the computer name can have.
<P><B>Warning:</B> If you skip the previous step, you may get back an empty string as the result. And you might crash the program if you fail to pre-size the buffer.
<P>
After that, we pass the pointer to the buffer and an integer value, which denotes the length of the buffer to the GetComputerName function. The strString buffer now contains the computer name and trailing spaces. In the final step we lop off the spaces to get the computer name, which we display, using the MsgBox function.<P>
(Of course an application that finds the Computer name has limited use, but here the focus was on understanding the mechanics of declaring and using an API function. So I deliberately chose a simple example.) <P>
<B>Conclusion</B><P>
In this article we discussed the mechanics of declaring and invoking an API function with the help of a simple example.

