# Automatically filter undone work in Excel.

## Background

In that time, I need to follow daily progress from vendor and the normal way that I need to do will be shown below

1. There are 3,000 works with 10 vendor in Excel file
2. 1 vendor owned about 300 works
3. I pick first work of vendor A then see current step. (assume now is step 5 of 11)
4. I ask vendor A that "What is current step ? is step 5 done ?"
5. Record the answer to Excel file.
6. Repeat 2,999 works left.

I solve all these problem by this VBA Excel. it can filter undone and lated works then send message to LINE app automaticcally.

## Requirement.

- Microsoft Excel
- Windows 10 or more (Only Windows can't be Macos) becuase there is no Thai language in VBA for Macos

## How to set up macro and use with Excel

### Step 1 : Paste personal.xlsb macro to its folder

1. Copy file `PERSONAL.xlsb` from project folder
2. Go to folder `C:\Users\{username}\AppData\Roaming\Microsoft\Excel\XLSTART`
3. Paste `PERSONAL.xlsb` into folder.

### Step 2 : Set up Thai language in Visual Basic

1. Open `HPB.xlsx` file.
2. Go to `Developer` tab (If there is no `Developer` tab, You need to enable it first).
3. Click at `Visual Basic` button.
4. On the left panel, expand `Chaperone (PERSONAL.XLSB)`.
5. Enter password as " " (spacebar 1 times).
6. Expand `Modules` 
7. Here you can see all VBA code but in wrong decode. You need to change language to Thai language.
8. Click `Tools` > `Options` at toolbar.
9. When `Option` window is pop up, Go to `Editor format` tab.
10. In font section, Choose any font which end with "(Thai)". For example, `CordiaUPC (Thai)`.
11. Click OK. Now language in code will display Thai language correctly.

### Step 3 : Edit LINE Notify Token to notify on LINE

1. Find function name `Progress_Button_Test()` in Modules2. This function for testing before send in real target.
2. Change to your LINE notify token of test group.
3. Find function name `Progress_Button_Real()` in Modules2. This function is real target.
4. Change to your LINE notify token of real group.
5. Close `Microsoft Visual Basic` windows.



## How to use

1. Select Vendor on A2 cell as you want of sheet "Report".
2. Select type of detail at B2 cell. There are 2 types.
   - Conclusion -> Just overall statistic.
   - Detail -> Detail for each progress of each work.
3. Select work phase at C2. There are 2 work phase.
- Front end -> Doing about job site outside office.
- Backend -> Doing about documentation within office.
4. Refresh work button at D2. To re-process and sumarize new data to table.
5. Send to test target at F2 button. | send to real target at E2 button.
6. Data in the left panel under controller is summary data. and table at the right is a detail of each process.