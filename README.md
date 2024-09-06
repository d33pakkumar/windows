# Retrieve Windows Product Key using VBScript

If other methods didn’t work, here’s an alternative approach using a VBScript to retrieve your Windows product key.

## Steps to Retrieve the Product Key

1. **Open Notepad** and paste the following code into it:

    ```vbscript
    Set WshShell = CreateObject("WScript.Shell")
    WScript.Echo ConvertToKey(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))

    Function ConvertToKey(Key)
        Const KeyOffset = 52
        i = 28
        Chars = "BCDFGHJKMPQRTVWXY2346789"
        Do
            Cur = 0
            x = 14
            Do
                Cur = Cur * 256
                Cur = Key(x + KeyOffset) + Cur
                Key(x + KeyOffset) = (Cur \ 24) And 255
                Cur = Cur Mod 24
                x = x - 1
            Loop While x >= 0
            i = i - 1
            KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
            If (((29 - i) Mod 6) = 0) And (i <> -1) Then
                i = i - 1
                KeyOutput = "-" & KeyOutput
            End If
        Loop While i >= 0
        ConvertToKey = KeyOutput
    End Function
    ```

2. **Save the file**:
   - Go to `File > Save As`, and in the "Save as type" dropdown, select **All Files**.
   - Name the file something like `GetWindowsKey.vbs`.

3. **Run the script**:
   - Double-click the `.vbs` file you just saved, and a popup will display your Windows product key.

## Notes

- This script reads the **DigitalProductId** from the Windows registry and converts it to the product key.
- The key may not be displayed if the Windows installation was activated through a digital license.

Let me know if this works for you!
