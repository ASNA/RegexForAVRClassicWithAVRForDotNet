## How to use an AVR for .NET DLL with AVR Classsic 5.0

AVR Classic is quite nearly a hermetically sealed environment. If you need AVR Classic to do something it doesn't do out the box you're only hope is finding, or creating, an ActiveX control or COM-capable DLL. Back in AVR Classic's halcyon days, there were hundreds of third-party ActiveX controls that added features such as FTP, regular expressions, DOS file access, user interface elements, and many others. Today, that marketplace is quite literally a ghost town. 

While it doesn't work for user interface enhancements, for pretty much any another enhancement you can think of that AVR Classic needs, ASNA Visual RPG for NET is a superb ally.  The other day a customer asked about how to implement regular expressions with AVR Classic. At first, we thought this was a pretty dark alley. We looked for ActiveX controls, but couldn't quickly find anything that seemed current and workable. 

Then it hit us. Like many ASNA customers, this customer has both AVR Classic and AVR for .NET. We dusted off our (considerably dusty!) COM interop skills and created a regex solution for AVR Classic with AVR for .NET. 

That solution isn't the subject of this article. Rather, this article focuses on the nuts and bolts of creating an AVR for .NET class library (which produces a DLL) and then registering that DLL for use with AVR Classic. An article is coming about the regex solution and you can [get a sneak peak at both its AVR for .NET and AVR Classic projects here](https://github.com/ASNA/RegexForAVRClassicWithAVRForDotNet). 

For now, let's focus on the general mechanics of using an AVR for .NET class library DLL with AVR Classic. While this article assumes the DLL is written with AVR for .NET, if the DLL doesn't need DataGate file access or IBM i program calls, the same instructions (nearly, there are some syntax differences) the DLL could also be written with C# or VB.NET.

> Beyond basic AVR Classic extensibility, if you're in the process of thinking about replacing your old AVR COM apps with AVR for .NET (and you should be!), the techniques presented here are a great place to start. You could do file IO, business logic, and solve lots of other business problems with AVR for .NET DLLS and use these DLLs in both the .NET and and COM worlds. 

#### Visual Studio setup 

Be we consider any code, we need to do a little setup so we the tools available to register .NET DLLs for COM use. To register a .NET DLL with COM, you need to use .NET's built-in RegAsm utility from the DOS command line. Before writing a DLL for COM's use, use these instructions to ensure you have easy access to RegAsm. 

>Keep in mind that you'll need to install the AVR for .NET runtime and the corresponding version of the .NET Framework on any end-user PCs for the DLLs to work. You'll also need to use RegAsm to register the .NET DLLs on the client PCs. 

Start Visual Studio and click Tools->Externals to bring up the dialog below. Click the 'Add' button and add the .NET Dev Prompt tool as shown below. 

![](https://asna.com/filebin/marketing/article-figures/vs-cmdprompt-tool.png?)

Title:

	.NET Dev Prompt
	
Command: 	

	C:\Windows\System32\cmd.exe
	
Arguments:

	/k "C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\Tools\VsDevCmd.bat"

Initial directory:

	$(BinDir) 

Set the checkboxes as shown in the image above. Click OK. 

Having added a new external tool, you can get to it directly from the Tools menu as shown below:

![](https://asna.com/filebin/marketing/article-figures/runvstool.png?)

When you run the .NET Dev Prompt, it will start in the project's `bin\debug` folder.  This is handy because that's where the DLL will be that you'll be registering. There are several different arguments available for specifying the initial folder. For use with other commands you might want the DOS command box to open at a different folder. [See the full list of VS external tool arguments here](https://msdn.microsoft.com/en-us/library/76712d27.aspx).

> The `Arguments` value shown above specifies the fully qualified path to the `vsdevcmd.cmd` file for Visual Studio 2015. This file is provided with every version of Visual Studio. If you're using a version of Visual Studio other than 2015, be sure to use the location of your version's `vsdevcmd.cmd` for the `Arguments` value. A way to find yours is to open a DOS command box and change the path to `c:\Program Files (x86)`  and then type `dir vsdevcmd.bat /s` and press enter. This will list all occurrences of that file and their parent folders. 

##### One more thing...

With the steps above, you've added a DOS command box that starts with the necessary path to use RegAsm (and other .NET tools, like ILDASM). However, when registering DLLs for COM use, RegAsm requires admin privileges. To ensure the DOS command box starts with admin privileges, stop Visual Studio and restart it by right-clicking it on the Windows menu and selecting "Run as administrator" (as shown below). When Visual Studio is run as an administrator, the DOS command box you run from Visual Studio will also be run as an administrator. Any time you need to register a DLL for COM, you'll need to start Visual Studio as an administrator.

![](https://asna.com/filebin/marketing/article-figures/![](https://asna.com/filebin/marketing/article-figures/runvsadadmin.png?)

> If you have Visual Studio pinned to your Windows taskbar, you can right click-its taskbar shortcut and use the context menu shown to use VS as an administrator.

#### Example code for a simple DLL for COM use

.NET DLLs intended to be used by COM applications must have their class declaration marked with the `ComVisible` and `ClassInterface` attributes and public methods must be marked with the `ComVisible` attribute as shown below. They must also have a single, explicitly declared constructor (with no parameters). No constructors overloads are allowed; there is no opportunity to pass parameters to the constructor in COM. It's OK to use a namespace with your DLL, but it won't be seen by COM. 

This example exposes a public method named Square that returns the square value of two integers. Be careful with the data types you return from functions used by COM. It's best to stick with AVR for .NET's *Integer4, *Packed, *Zoned, and *String data types. The line continuations below are to keep the code narrow for publication purposes.

    Using System.Runtime.InteropServices
    Using System.Text.RegularExpressions
    
    BegClass Math Access(*Public)  +
      Attributes(ComVisible(*True), +
        ClassInterface(ClassInterfaceType.AutoDual))
    
        BegConstructor Access(*Public) 
        EndConstructor 
    
        BegFunc Square Type(*Integer4) Access(*Public) +
          Attributes(ComVisible(*True)) 
            DclSrParm Value1 Type(*Integer4) 
            DclSrParm Value2 Type(*Integer4) 
    
            LeaveSr (Value1 * Value2)
        EndFunc 
    EndClass        

Do all you can to keep the classes you create for COM consumption as simple as possible. Events, overloaded functions, returning complex types from functions, and pretty much anything beyond returning simple types will lead to problems. Keep it simple, Sam! Also, be sure to thoroughly test your .NET libraries with .NET code before you try to use them in AVR Classic. Don't attempt to use a .NET DLL with AVR Classic until you are sure the DLL works as expected. You can't easily debug .NET code when called from from AVR Classic.

##### Change AssemblyInfo

In the project's AssemblyInfo.vr file (which is under your projectc's Properties folder in the Solution Explorer), you must also change the ComVisible() attribute value from *False to *True:

Change this line:

	SetAssemblyAttribute ComVisible(*False)

to this:

	SetAssemblyAttribute ComVisible(*True)

Compile the DLL with Build->Build Solution.

#### Registering/unregistering with COM

To register a .NET DLL for COM use, you need to use .NET's RegAsm utility. This utility needs to be in your path and run from a DOS command box with administrative privileges. See the "Visual Studio setup" above for more information. 
    
To register the DLL with COM (where `mydll.dll` is the of your DLL):

	regasm /tlb /codebase mydll.dll

To unregister the DLL with COM:
    
    regasm /unregister mydll.dll

For this example, the Visual Studio project as named ComTest and its resulting DLL was named ComTest.DLL. 

#### Using the .NET DLL from AVR Classic 

##### Setting a reference to your .NET DLL

From the AVR Classic project from which you want to use your .NET DLL, use the Project->References to add a reference to the your DLL. In this case, we're setting a reference to the `ComTest.DLL` by checking its checkbox. Click `OK`.

![](https://asna.com/filebin/marketing/article-figures/avr-classic-references.png)

##### Confirm the DLL and its classes with AVR Classic's Object Browser

After setting the reference to your DLL, use AVR Classic's Object Browser to confirm it is loaded. Use View->Object Browser to start the object browser, then select your library, which should be the name of your DLL without the .DLL extension) from the dropdown. 

You should see your library's classes and their members displayed. You'll notice that three methods (Equals, GetHashCode, and GetType) are displayed along with the methods your provided. These are standard methods every .NET class gets--and you're not likely to use them in the COM project. 

![](https://asna.com/filebin/marketing/article-figures/avr-classic-object-browser.png)

When you select a class in the object browser, its footer area shows you how to declare that class in your AVR Classic code. In this case, the Math class is declared like this: 

	DclFld x Type(ComTest.Math)

##### Using the .NET class in your AVR code

At this point, you've done the hard work, all that remains is the simple code need to declare your .NET class and then put it to work. Note for AVR for .NET coders, unlike .NET, COM implicitly instances the class--you don't use the New keyword to instance it. 

	DclFld ComTestLib Type(ComTest.Math) 
	DclFld sqr Type(*Integer) Len(4) 

	sqr = ComTestLib.Square(4,4) //sqr will be 16!

#### Changing the .NET DLL

If you need to change the code for your .NET DLL:

* First, end AVR Classic. This is very important. If you make changes to a DLL that an AVR Classic project is using,  those changes will not be seen until you end and restart AVR Classic. 
* Unregister the DLL with: `regasm /unregister <DLL name>` in a DOS command box started with administrator privileges  and with `RegAsm` in its path. Before sure to do this before you make changes to the DLL's code.
* Make changes to the .NET code and recompile the DLL. Reregister it with `regasm /tlb /codebase <DLL name>` in a DOS command box started with administrator privileges  and with `RegAsm` in its path. 
* Start AVR Classic again and load your project. It should recognize the changes.
