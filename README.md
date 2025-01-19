# Uncommon VBScript Errors

This repository demonstrates two uncommon error scenarios in VBScript:

1. **Implicit Type Coercion:** VBScript's flexible type system can lead to unexpected behavior when comparing different data types. The code example shows how a string comparison with a number may seem to work, but is unreliable.
2. **Range Check Edge Cases:**  A common function to check if a number falls within a range often overlooks edge cases where the number is equal to the minimum or maximum of the range. The example demonstrates how a simple range check function can fail to handle equality correctly.

The `bug.vbs` file contains code exhibiting these issues, while `bugSolution.vbs` offers improved, more robust versions.