# One-Click Leadsheet Generator (OLSG)

[![Watch the video](images/youtube_screenshot.png)](https://www.youtube.com/watch?v=S0zmJipumgw)


OLSG is an Excel Add-in which enables the auditor to easily generate 
[lead schedules](https://auditnz.parliament.nz/resources/working-with-your-auditor/csf/lead-schedules) (leadsheets) 
required for auditing purposes, using the Trial Balance (TB) as input. 

The only requirement on the auditor side is to prepare the TB in the appropriate format. 
The auditor determines which accounts are significant and then the leadsheets are just one click away (it is 
literally one click). 

OLSG creates separate leadsheets (.xlsx format) for each of the accounts that were determined as significant by the 
auditor. For the accounts that were not determined as significant, OLSG generates a file named “non-significant 
leadsheets” (.xlsx format) which presents the leadsheets for the rest of the accounts in separate tabs.

The created leadsheets are then used by the auditor to perform the audit procedures.



## Installation
- Clone the repository locally using `git clone`
- Pip install the requirements
- Install the [xlwings addin](https://docs.xlwings.org/en/stable/addin.html) in Excel: `xlwings addin install`
(if you encounter permission issues, see the Contribution section later on how to provide admin rights to Pycharm)
- Create your leadsheet in the expected format (see `demo trial balance.xlsx` for an example)
- Make sure xlwings is added as an addin in Excel
- Press the Run button and let Leadsheet Generator do its magic!

### Compatibility

Currently, the tool is fully functional in the latest Excel version for Mac.

## Contribution

To run the test with Pycharm, make sure that Pycharm has admin rights.

To start Pycharm (Community Edition) with admin rights in Mac, use:

```commandline
alias pycharm="/Applications/PyCharm\ CE.app/Contents/MacOS/pycharm"

pycharm
```
