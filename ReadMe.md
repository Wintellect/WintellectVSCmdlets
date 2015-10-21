# Wintellect VS Cmdlets Module #

With Visual Studio dropping support for macros way back in VS 2012, the only way to do simple customizations is to download the VS SDK, write a VSIX, and have VS debugging VS. Personally, I find it ridiculous that we can no longer do simple extensibility in our main tool.

Because I've written numerous small macros, and use them on a daily basis, I was not looking forward to the extra effort to turn them into a VSIX. Fortunately, Visual Studio has the wonderful NuGet package manager built in so I converted my macros into PowerShell. While not as convenient as macros, they do work. You can read more about converting macros into PowerShell on my [blog ](http://www.wintellect.com/devcenter/jrobbins/using-nuget-powershell-to-replace-missing-macros-in-dev-11).

These macros/cmdlets work for Visual Studio 2010 through Visual Studio 2015.Please fork and let me know if there's any bugs you find. I hope you find them useful.

Here's the about text showing all cmdlets. Of course, all cmdlets have detailed help for more information.

	TOPIC
	    about_WintellectVSCmdlets
	    
	SHORT DESCRIPTION
	    Provides missing functionality, especially around debugging, to Visual Studio 2010 and Visual Studio 2012.
	           
	LONG DESCRIPTION
	    This describes the basic commands included in the WintellectVSCmdlets module. With VS 2012 not offering
	    macros, simple extensions require installing an SDK and debugging the extensions with a second instance
	    of the IDE. In all, it makes for a very poor experience when you want to do simple customization of the 
	    the development environment.
	    
	    These macros, which are very useful for debugging, demostrate that the NuGet Package Console is 
	    sufficient for many of your customization needs. Most of these cmdlets are ports of VB.NET macros that 
	    John Robbins has shown on his blog and books.
	    
	    All cmdlets work with Visual Studio 2010 through Visual Studio 2015.
	    
	    Note that these cmdlets support C#, VB, and native C++. They probably support more but those were
	    all the languages tested.
	
	    If you have any questions, suggestions, or bug reports, please contact John at john@wintellect.com.
	                 
	    The following Wintellect VS cmdlets are included.
	
	        Cmdlet					                        Description
	        ------------------		                        ----------------------------------------------
	        Add-BreakpointsOnAllDocMethods                  Sets breakpoints on methods in the current code document. This
	                                                        is very useful in .NET languages as the debugger expression 
	                                                        evaluator does not support that.
	                                                        
	        Remove-BreakpointsOnAllDocMethods               Removes all the breakpoints set with Add-BreakpointsOnAllDocMethods.
	                                                        This cmdlet will not remove any of your breakpoints.
	        
	        Add-InterestingThreadFilterToBreakpoints        Adds the filter "ThreadName==InterestingThread" to all breakpoints to
	                                                        make it easier to debug through a single transaction.
	                                                        
	        Remove-InterestingThreadFilterFromBreakpoints   Removes the "ThreadName==InterestingThread" filter applied with
	                                                        Add-InterestingThreadFilterToBreakpoints.
	
	        Disable-NonActiveThreads                        Freezes all but the active thread so you can single step to the end
	                                                        of a method without dramatically bouncing to another thread when you 
	                                                        least expect it.
	                                                        
	        Resume-NonActiveThreads                         Thaws all threads previously frozen with Disable-NonActiveThreads.
	
	        Get-Breakpoints                                 Returns the latest version of the IBreakpoints derived list.
	        
	        Get-Threads                                     Returns all the threads.
	        
	        Invoke-NamedParameter                           A wonderful cmdlet that lets you easily call methods with many
	                                                        optional parameters. Full credit to Jason Archer for this cmdlet.
	        
            Invoke-WinDBG                                   Visual Studio has ease of use, where WinDBG (with SOS + SOSEX) have 
                                                            tons of power to tell you what's going on in your application. This
                                                            cmdlet starts WinDBG on the process you're currently debugging in the
                                                            IDE so you can have the best of both worlds.

	                                                        
	        Open-LastIntelliTraceRecording                  When you stop debugging, your current IntelliTrace log disappears. This
	                                                        cmdlet fixes that by opening the last log you produced so you can post-mortem
	                                                        look at your debugging session.
	
	SEE ALSO
	    Online help and updates: http://www.wintellect.com/devcenter/author/jrobbins
	    Add-BreakpointsOnAllDocMethods
	    Remove-BreakpointsOnAllDocMethods
	    Add-InterestingThreadFilterToBreakpoints
	    Remove-InterestingThreadFilterFromBreakpoints
	    Disable-NonActiveThreads
	    Resume-NonActiveThreads
	    Get-Breakpoints
	    Get-Threads
	    Invoke-NamedParameter
	    Invoke-WinDBG
	    Open-LastIntelliTraceRecording
	
