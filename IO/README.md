# IO System (Input/Output) `IOEngine.bas`

**Author:** RokyBeast\@RokyXxX
**License:** MIT
**Language:** VBA (PowerPoint Compatible)

---

## 🔍 Overview

This module implements a simple and reusable Input/Output system in VBA, primarily designed to handle keyboard input using Windows API calls. It can be used in interactive PowerPoint projects such as custom UI systems or games.

The system runs a loop that checks for key states using `GetAsyncKeyState` and performs actions accordingly.

---

## 📚 Key Features

* Compatible with both **VBA7** and **pre-VBA7** environments.
* Uses Windows API `GetAsyncKeyState` for real-time key detection.
* Custom polling delay via `delayMs`.
* Toggleable input loop using `InitIO` and `EndIO`.

---

## 📦 Public Interface

### `Public vRunning As Boolean`

* Tracks the current state of the input loop.

### `Public Const delayMs As Integer = 10`

* Specifies the delay in milliseconds between each polling cycle.

### `Public Property Let IOState(NewIO As Boolean)`

* Setter to start or stop the polling loop.

---

## 🧩 Functions & Subs

### `InitIO()`

* **Purpose**: Starts the input handling loop.
* **Internals**: Sets `vRunning` to `True` and calls `IOHandler()`.

### `EndIO()`

* **Purpose**: Stops the input handling loop.
* **Internals**: Sets `vRunning` to `False`.

### `IOHandler()`

* **Purpose**: Main loop that continuously checks for a key press.
* **Note**: Replace `KEY_NAME` with a valid key constant.
* **Mechanism**:

```vba
If GetAsyncKeyState(KEY_NAME) And &H8000 Then
    ' Handle Key Press Event
End If
```

### `Delay(Optional ms As Long = 10)`

* **Purpose**: Creates a non-blocking delay.
* **Implementation**: Uses `Timer` and `DoEvents`.

---

## 🧠 Usage Example

```vba
Sub ExampleUsage()
    InitIO
    ' Let it run, then stop
    Delay 3000 ' Let it run for 3 seconds
    EndIO
End Sub
```

Replace `KEY_NAME` with constants like `vbKeyA`, `vbKeySpace`, etc. Refer to [KeyCode Constants](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/keycode-constants) for more.

---

## 🔐 API Declaration

```vba
#If VBA7 Then
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If
```

---

## ✅ Notes

* Ensure your macro settings allow Windows API calls.
* Use `DoEvents` wisely to avoid freezing PowerPoint.
* This is a non-blocking, passive polling system.

---

## 📄 License

MIT License

---
