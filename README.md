# Math-Interpreter-VBA
## Description
This project is a VBA Math Interpreter, made without any external dependencies. <br/> It supports basic operations, floats numbers, negative numbers, operation with brackets and unary expressions. <br/> (This is my first interpreter on VBA yay!)

## How to use
First you have to download the module, and import it on PowerPoint or Excel. <br/>
After that you should be able to use the function ``Evaluate`` to calculate.

```vb
' It returns the result in a string.
Eval.Evaluate(Expression)   
```

## Documentation
> Arithmetic Operators

| Operator | Name | Priority |
| --- | --- | --- |
| ^ | Power | 3 |
| * | Multiply | 2 |
| / | Division | 2 |
| + | Sum | 1 |
| - | Subtraction | 1 |

> Unary Expressions

| Examples | Name |
| --- | --- |
| ++ | Positive | 
| -- | Positive | 
| +- | Negative | 
| -+ | Negative | 
| --- | Negative | 

