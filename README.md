<div align="center">

## To Keep a form on top of another \(Only on top of one form, not of the whole application\)


</div>

### Description

This is a very simple and precise solution to keep a form on top of only one other form. I was looking for a solution, but couldn't find it anywhere. All the codes and suggestions either keep the form on top of All windows programs or on top of the whole application. But this one will only keep one form on top of another one.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Faisal Nawaz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/faisal-nawaz.md)
**Level**          |Beginner
**User Rating**    |4.3 (30 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/faisal-nawaz-to-keep-a-form-on-top-of-another-only-on-top-of-one-form-not-of-the-whole-app__1-44448/archive/master.zip)





### Source Code

```
' Let say you have two forms in your project, i.e. form1.frm and form2.frm. Now you want to open form2.frm from form1, but want your form2 to stay on top of form1 and also want to access the menu of form1 without causing the form2 to go behind. Here is the simple solution...
=========================================
Private Sub Command1_Click()
   Form2.Show vbModeless, Form1
End Sub
=========================================
I hope it will help you in your projects...
```

