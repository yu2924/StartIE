# StartIE

 ## What is this?

This small piece of code launches traditional Internet Explorer using the traditional Win32 API, even on Windows 11 with IE disabled.  

In fact, this code is equivalent to the following script, which simply activates the application through the OLE automation interface:
```
app = new ActiveXObject("InternetExplorer.Application");
app.Visible = true;
```


## Requirement
* Windows SDK
* ATL

## Usage
Just run from the command or double click the icon.  
no parameters required.

## written by
[yu2924](https://twitter.com/yu2924)

## License
CC0 1.0 Universal
