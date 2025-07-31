# Math-Interpreter-VBA
## Description
A **fully functional mathematical expression interpreter** built entirely in **VBA (Visual Basic for Applications)**.  
It parses and evaluates math expressions with **correct precedence**, **associativity**, and **unary operator support**,  
making it ideal for **Excel macros**, **Office automation**, or just exploring compiler logic in VBA.

## How to use
First you have to download the module, and import it on PowerPoint or Excel. <br/>
After that you should be able to use the function ``Evaluate`` to calculate an expression in a String.

```vb
' It returns the result in a string.
Eval.Evaluate(Expression)   
```
## Features

✅ Built from scratch using the **Shunting Yard algorithm**  
Parses infix expressions into postfix (RPN)  
Evaluates expressions with:

| Supported | Description                    |
|-----------|--------------------------------|
| `+ - * / ^` | Arithmetic operators           |
| `()`       | Parentheses for grouping       |
| `-x`, `+x` | Unary operators (e.g., `-5`)   |
| `1.5`      | Decimal numbers                |

