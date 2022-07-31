# RewardsTool
An application for Windows written in C# with the [Selenium library](https://www.selenium.dev/) and that resolves automatically quizzes and researches on [Microsoft rewards](https://account.microsoft.com/rewards/) to earn points.  


## Requirements
  - ~~Microsoft Edge based on Chromium~~ Chrome

## Dependencies
  - [Selenium.WebDriver 4.3.0](https://www.nuget.org/packages/Selenium.WebDriver/4.3.0?_src=template)
  - ~~[Microsoft.Edge.SeleniumTools](https://www.nuget.org/packages/Microsoft.Edge.SeleniumTools/3.141.2?_src=template)~~
  - [DotNetSeleniumExtras.WaitHelpers](https://www.nuget.org/packages/DotNetSeleniumExtras.WaitHelpers/3.11.0?_src=template)
  - [System.IO.Compression](https://www.nuget.org/packages/System.IO.Compression/4.3.0?_src=template)
  - [System.IO.Compression.ZipFile](https://www.nuget.org/packages/System.IO.Compression.ZipFile/4.3.0?_src=template)

## Arguments (optional)

  - **-w**
    - The program will not start until the user will not press a button, useful if the program is executed each windows startup so in case you can stop it before starting
    

  - **-p Path\of\folder**
    - The program will install Chrome driver in the selected folder, you can put your driver in the folder if you want and the program will try to use it, but it will not be enabled to update the driver unleast you give the permission.  
    If not specified the programm will use the folder where is located


- **-u Profile**
    - The program will open Chrome using the specified profile. If not specified the program will use the default profile
    
## Errors
It may occur that an error interrupts the program, in this case close Chrome and restart the program. If an error also occurs in second attempt [report it](https://github.com/Metgraph/RewardsTool/issues).

After reported the error you can try to complete the cards manually and check if the next day the problem reoccurs (in this case the problem is correlated with a single card)

## Supported nation
Italy

