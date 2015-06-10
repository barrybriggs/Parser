## Parser ##

**Parser** is a command line tool that simulates parsing and evaluating Excel-like formulas. I use it as a test harness for my cloud-based computation engine codenamed "CloudSheet." 

Parser is a recursive descent style parser and a basic evaluator (with a lot of stubs) is included so things like 
    
    =sum(3,5)

will work, as will, hopefully, much more complex formulas. 

However, as a test harness, it's a work in progress and I suspect some things may have gotten fixed in the non-test version that didn't get reflected back to this version. 

To build and run, simply create a Visual Studio console application and copy this repo's `program.cs` file over the default. An overview of the design of the Parser is located on this [blog](http://blogs.msdn.com/partnercatalystteam), or will be very soon.  