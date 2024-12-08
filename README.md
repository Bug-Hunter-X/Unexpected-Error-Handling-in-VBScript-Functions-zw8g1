# Unexpected Error Handling in VBScript Functions
This example demonstrates a subtle error in VBScript's error handling.  When calling a function, if required parameters are omitted (passed as empty), the script might not raise the intended error. This can cause unexpected crashes or incorrect results.

## Bug
The `bug.vbs` file shows a function that checks for empty parameters. However, due to the way VBScript handles errors, if the function is called without parameters, it might not always raise the error correctly.

## Solution
The `bugSolution.vbs` file provides a more robust solution. It uses a more explicit check for empty parameters which ensures the expected error will be raised.