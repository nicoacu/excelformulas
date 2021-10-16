# Language
- [EN](#en)
- [ES](#es)
#

### [EN]
# Excel/Google Sheets Formulas

My frequent go-to formulas and their respectives uses. To clarify, some formulas can be applied in both systems (excel/google sheets) but might require some changes to make it work.


## Content

- [Excel Formulas](#excel-formulas)
  - [Content](#content)
      - [Text Manipulation](#text-manipulation)
        - [Find if a cell contains an empty space](#find-if-a-cell-contains-an-empty-space)
        - [Search for a value using two columns as reference](#search-for-a-value-using-two-columns-as-reference)
        - [Remove empty spaces from a cell](#remove-empty-spaces-from-a-cell)
        - [Search for a keyword within a cell and categorize it with a label](#search-for-a-keyword-within-a-cell-and-categorize-it-with-a-label)
        - [Replace an #ERROR with some text](#replace-an-error-with-some-text)
     - [Date formats](#date-formats)
        - [Cheatsheet to convert to different type of formats](#cheatsheet-to-convert-to-different-type-of-formats) 
     - [Number formats](#number-formats)
        - [Show 1K and 1M instead of 1,000 and 1,000,000](#show-1k-and-1m-instead-of-1000-and-1000000)  

#

- [Google Sheets Formulas](#google-sheets-formulas)
     - [Array Formulas](#array-formulas)
       - [Fill range with formulas using another cell as a reference](#fill-range-with-formulas-using-another-cell-as-a-reference)
     - [Data Import](#data-import)
       - [Import filtered data from another Google Sheet](#import-filtered-data-from-another-google-sheet)    

#

- [Bonus tricks and shortcuts](#bonus-tricks-and-shortcuts)
  - [Check the value of a cell inside the formula bar](#check-the-value-of-a-cell-inside-the-formula-bar)
  - [Shortcut to create new google sheet/doc/slide](#shortcut-to-create-new-google-sheetdocslide)
  - [Shortcut to Find and Replace in Google Sheets](#shortcut-to-find-and-replace-in-google-sheets) 


#
# Excel Formulas
### Text Manipulation

#### Find if a cell contains an empty space

Let's say that you have a range of cells where some of them are not entirely empty but are completed with an empty space. This might impact you if you are using these cells for other formulas/purposes, as their data type are not the same [(Read: difference between null and an empty space string)](https://www.mrexcel.com/board/threads/null-value-vs-empty-cell-vs-vs-0-vs-blank-cell.468838/). There are several ways to check if a cell is actually empty, but this is the one that i would use to check that:       

``` bash
=IF(COUNTBLANK(A1),"is empty","has a blank space")
```
If by any chance you would need the IF statement to tell you which is the cell that was checked, you could also use this formula below, modifying only the "A1" part of it to the cell that you need to check.

``` bash
=IF(COUNTBLANK(A1),SUBSTITUTE(CELL("address",A1),"$","")&" is empty",SUBSTITUTE(CELL("address",A1),"$","")& " has a blank space")
```

<i>Expected output:</i> `A1 is empty, A1 has a blank space`

#

#### Search for a value using two columns as reference

There are some situations where you might need to search for a value using more than one column as a reference. In this case, I would use an Array Index-Match formula [(Read: Create an Array Formula in Excel)](https://support.microsoft.com/en-us/office/create-an-array-formula-e43e12e0-afc6-4a12-bc7f-48361075954d) which allows us to include more than one range of cells as a reference, but before doing this, is important to understand the sintaxis of a normal Index Match formula, which is explained [here](https://support.google.com/docs/answer/3098242?hl=en). The main advantage of using an Index-Match formula instead of a vlookup is the flexibility that it provides (as VLOOKUP can only be used when the lookup value is to the left of the desired attribute to return).

Let's use this hypothetical table as an example:

<table>
<tr>
  <th> </th>
  <th>A</th>
  <th>B</th>
  <th>C</th>
</tr>
<tr>
  <th>1</th>  
  <td>Animal</td>
  <td>Age</td>
  <td>Name</td>
</tr>
<tr>
  <th>2</th>    
  <td>Cat</td>
  <td>2</td>
  <td>Missy</td>
</tr>
<tr>
  <th>3</th>  
  <td>Dog</td>
  <td>5</td>
  <td>Confetti</td>
</tr>
<tr>
  <th>4</th>  
  <td>Dog</td>
  <td>4</td>
  <td>Coco</td>
</tr>
<tr>
  <th>5</th>  
  <td>Cat</td>
  <td>4</td>
  <td>Simba</td>
</tr>
<tr>
  <th>6</th>  
  <td>Cat</td>
  <td>3</td>
  <td>Kyra</td>
</tr>
</table>

If we would use an index match formula to return the name of a dog, we would use

``` bash
=INDEX($A$1:$C$6, MATCH("dog",$A$1:$A$6,0),2)
```

<i>Expected output:</i> `Confetti`, since its the first dog name that appears on the table. 

Now if we wanted to know which is the name of a 4 y/o Cat, we would need to use an array index match formula:

``` bash
=INDEX($A$1:$C$6, MATCH("cat"&"4",$A$1:$A$6&$B$1:$B$6,0),2)
```

<i>Expected output:</i> `Simba`

<b>Note: this formula will only work in excel if you press [CTRL + SHIFT + ENTER](https://support.microsoft.com/en-us/office/create-an-array-formula-e43e12e0-afc6-4a12-bc7f-48361075954d), otherwise i'll throw a #VALUE error.</b> 

#

#### Remove empty spaces from a cell

<i>Example:</i> `The table is over the cat`

``` bash
=SUBSTITUTE("the table is over the cat"," ","")
```
<i>Expected output:</i> `Thetableisoverthecat`

If for some reason you would need to include nonbreaking space characters, you could also use:

``` bash
=SUBSTITUTE(A1,CHAR(160),"")
```

#

#### Search for a keyword within a cell and categorize it with a label

There are several ways to create a 'Category' column. In this example I'll show one that helped me on several occassions. 

Let's use this table as an example:

<table>
<tr>
  <th> </th>
  <th>A</th>
  <th>B</th>
</tr>
<tr>
  <th>1</th>  
  <td>Category</td>
  <td>Product</td>
</tr>
<tr>
  <th>2</th>    
  <td> </td>
  <td>Cellphone</td>
</tr>
<tr>
  <th>3</th>  
  <td> </td>
  <td>Computer</td>
</tr>
<tr>
  <th>4</th>  
  <td> </td>
  <td>Computer Desk</td>
</tr>
<tr>
  <th>5</th>  
  <td> </td>
  <td>Chair</td>
</tr>
<tr>
  <th>6</th>  
  <td> </td>
  <td>Table</td>
</tr>
</table>

Let's imagine that we have a list of items and we would like to categorize these as 'Technology', 'Furniture' . We could use a nested IF formula similar to this one to do it:

``` bash
=IF(OR(B4="Cellphone",B4="Computer"), "Technology", IF(OR(B4="Desk",B4="Chair",B4="Table"), "Furniture", "Uncategorized"))
```

However, that formula would bring us issues to categorize "Desk", as the value in the cell B4 says "Computer Desk" and not only "Desk". To fix this, we could rearrange the formula a bit and use this one instead

``` bash
=IF(OR(B4="Cellphone",B4="Computer"), "Technology", IF(OR(ISNUMBER(SEARCH("Desk",B4)),B4="Chair",B4="Table"), "Furniture", "Uncategorized"))
```

<i>Expected output:</i> 

<table>
<tr>
  <th> </th>
  <th>A</th>
  <th>B</th>
</tr>
<tr>
  <th>1</th>  
  <td>Category</td>
  <td>Product</td>
</tr>
<tr>
  <th>2</th>    
  <td>Technology</td>
  <td>Cellphone</td>
</tr>
<tr>
  <th>3</th>  
  <td>Technology</td>
  <td>Computer</td>
</tr>
<tr>
  <th>4</th>  
  <td>Furniture</td>
  <td>Computer Desk</td>
</tr>
<tr>
  <th>5</th>  
  <td>Furniture</td>
  <td>Chair</td>
</tr>
<tr>
  <th>6</th>  
  <td>Furniture</td>
  <td>Table</td>
</tr>
</table>

What the =ISNUMBER(SEARCH("Desk",B4)) formula does is to search for the position of the word "Desk" inside the string "Computer Desk", and if that formula throws a number (which it would do, if the word is in the string of B4) the =ISNUMBER() would return a TRUE boolean result, and since this is all inside an =OR() formula (which basically executes the TRUE condition of the IF formula, if any of the checks inside the =OR() throw a TRUE result) it ends up returning the category that we wanted, which is 'Furniture'.      


Again, this is one of many ways to label a keyword. in Google Sheets you could even use =IFS() to perform a similar function to a nested IF but more legible.


#

#### Replace an #ERROR with some text

A generic way to replace an #ERROR with a specific text would be to use the =IFERROR formula at the start of your original formula. For Example:

``` bash
=IFERROR(INDEX($A$1:$C$6, MATCH("dog",$A$1:$A$6,2)),"This is giving me an error")
```

However you can replace some specific errors instead if you need to (e.g: replacing a #N/A error) 

``` bash
=IF(ISNA(INDEX($A$1:$C$6, MATCH("dog",$A$1:$A$6,0),2)),"This is a #NA error", INDEX($A$1:$C$6, MATCH("dog",$A$1:$A$6,0),2))
```

#
### Date Formats

#### Cheatsheet to convert to different type of formats

<table>
<tr>
  <th></th>
  <th>A</th>
  <th>B</th>
  <th>C</th>
</tr>
<tr>
  <th>1</th>  
  <td>Current Format</td>
  <td>Wanted Format</td>
  <td>Formula</td>
</tr>
<tr>
  <th>2</th>    
  <td>10/31/2021</td>
  <td>October 31, 2021</td>
  <td>=TEXT(A2,"mmmm dd, yyyy")</td>
</tr>
<tr>
  <th>3</th>  
  <td>10/31/2021</td>
  <td>Sunday 31 October, 2021</td>
  <td>=TEXT(A3, "dddd d mmmm, yyyy")</td>
</tr>
  <tr>
  <th>4</th>  
  <td>10/31/2021</td>
  <td>31/10/2021</td>
  <td>=TEXT(A3, "dd/mm/yyyy")</td>
</tr>
</table>

You might noticed that I use a =TEXT formula to change these formats, keep in mind that using that formula you are converting the data from the cell to a string, so, for example, if you are suming two cells with a date inside (which you can do), you won't be able to do so with the output of the =TEXT formula, as these values won't be considered numbers but will be considered text.

#
### Number formats

#### Show 1K and 1M instead of 1,000 and 1,000,000

To do this you need to apply a custom number format in the selected cells. If you need to change the format from 1,000 to 1K, use the following custom number format:

``` bash
0.0, “K”
```

If instead, for example, you need to change 1,000,000 to 1M, use the following custon number format:

``` bash
0.0,, “M”
```

Note that these format modifications are only changing the way that the numbers looks in the sheet, but the numbers stored in the cell are the same. If you want to change the numbers from the cell to a different format (e.g: from number format to text format) you can use the =TEXT formula mentioned before. example: 

``` bash
=TEXT(1000,"#,##0,K")
```

<i>Expected output:</i> `1K`

#
# Google Sheets Formulas
### Array Formulas

#### Fill range with formulas using another cell as a reference

Have you ever had a situation where you need to extend the range of your formulas to adapt them to the range of your current data set? Did you ever think about how annoying is to do this manually, specially if you need to update your data set frequently? well, you could use an =INDIRECT formula and if you are creative with it you can readapt ranges in a creative way. for example:

``` bash
=SUM(INDIRECT("H26:H")&ROW(A296)-1))
```
Would basically translate the formula to this

<i>Expected output:</i> `=SUM(H26:H295)`

But since it might be complex to use the INDIRECT formula to refer to ranges, there is another way to do this with the use of =ARRAYFORMULA(). <i>Note: using a lot of arrayformulas in a file can make it work slowly or even make it more prone to crash, so keep that in mind in case that you want to implement it for this use.</i> 

=ARRAYFORMULA() is, basically, a single formula in a cell with an output that can be multiples rows/columns. For example purposes, let's check the following table:

<table>
<tr>
  <th></th>
  <th>A</th>
  <th>B</th>
  <th>C</th>
</tr>
<tr>
  <th>1</th>  
  <td>Name</td>
  <td>Surname</td>
  <td>Complete Name</td>
</tr>
<tr>
  <th>2</th>    
  <td>Geralt</td>
  <td>Of Rivia</td>
  <td> </td>
</tr>
<tr>
  <th>3</th>  
  <td>Yennefer</td>
  <td>Vengerberg</td>
  <td> </td>
</tr>
  <tr>
  <th>4</th>  
  <td>Cirilla Fiona</td>
  <td>Elen Riannon</td>
  <td> </td>
</tr>
</table>

If we needed to complete column "C" with a formula, we could simply concatenate cells in column A and B to do this. for example, we could use =A2&" "&B2 and the output would be this: 

<table>
<tr>
  <th></th>
  <th>A</th>
  <th>B</th>
  <th>C</th>
</tr>
<tr>
  <th>1</th>  
  <td>Name</td>
  <td>Surname</td>
  <td>Complete Name</td>
</tr>
<tr>
  <th>2</th>    
  <td>Geralt</td>
  <td>Of Rivia</td>
  <td><b>=A2&" "&B2</b></td>
</tr>
<tr>
  <th>3</th>  
  <td>Yennefer</td>
  <td>Vengerberg</td>
  <td> </td>
</tr>
  <tr>
  <th>4</th>  
  <td>Cirilla Fiona</td>
  <td>Elen Riannon</td>
  <td> </td>
</tr>
</table>

  <i>Expected output for C2:</i> `Geralt of Rivia`
  
However we would need to drag down the formula in C3 and C4 to do the same in those fields. In this case that's easy to do, but as mentioned before, if your table changes frequenly and you don't want to manually drag that down, you could use in the cell C2 of this example =ARRAYFORMULA(A2:A&" "&B2:B) and that would automatically fill the rest!
  
<table>
<tr>
  <th></th>
  <th>A</th>
  <th>B</th>
  <th>C</th>
</tr>
<tr>
  <th>1</th>  
  <td>Name</td>
  <td>Surname</td>
  <td>Complete Name</td>
</tr>
<tr>
  <th>2</th>    
  <td>Geralt</td>
  <td>Of Rivia</td>
  <td><b>=ARRAYFORMULA(A2:A&" "&B2:B)</b></td>
</tr>
<tr>
  <th>3</th>  
  <td>Yennefer</td>
  <td>Vengerberg</td>
  <td> </td>
</tr>
  <tr>
  <th>4</th>  
  <td>Cirilla Fiona</td>
  <td>Elen Riannon</td>
  <td> </td>
</tr>
</table>  

<i>Expected output:</i>

<table>
<tr>
  <th></th>
  <th>A</th>
  <th>B</th>
  <th>C</th>
</tr>
<tr>
  <th>1</th>  
  <td>Name</td>
  <td>Surname</td>
  <td>Complete Name</td>
</tr>
<tr>
  <th>2</th>    
  <td>Geralt</td>
  <td>Of Rivia</td>
  <td><b>Geralt of Rivia</b></td>
</tr>
<tr>
  <th>3</th>  
  <td>Yennefer</td>
  <td>Vengerberg</td>
  <td><b>Yennefer Vergerberg</b></td>
</tr>
  <tr>
  <th>4</th>  
  <td>Cirilla Fiona</td>
  <td>Elen Riannon</td>
  <td><b>Cirilla Fiona Elen Riannon</b></td>
</tr>
</table>  

Here is a recommended [video](https://www.youtube.com/watch?v=ndxOEbgeqoQ), if you need more examples to understand how to use =ARRAYFORMULA 


#
### Data Import
#### Import filtered data from another Google Sheet

There is an easy and fantastic formula to import data from one google sheet to another one, and that is =IMPORTRANGE("spreadsheet_url", "range_string"). As a reference, here is the [link](https://support.google.com/docs/answer/3093340?hl=en) to the documentation page in case that you need it.

There is also another fantastic formula called =QUERY("data", query, [headers]) that works similar to a query in SQL. I'll also leave here the [link](https://support.google.com/docs/answer/3093343?hl=en) to the documentation page for reference.

Now you know what's the best part? you can use both: a =QUERY(=IMPORTRANGE(), query, [headers])) formula to bring information from one tab (protip: you can even use more than one importrange formula to bring information from more than one tab) filtered with the query needed. It's hard to showcase this with Github tables, so I recommend to check this [video](https://www.youtube.com/watch?v=Jl964i0UuCY) instead if you are interested in learning how to do this.


#
# Bonus tricks and shortcuts

#### Check the value of a cell inside the formula bar

This works for both Excel and Google Sheets, but i would recommend to test it in Excel as it's more evident.

Let's use this table as an example:

<table>
<tr>
  <th> </th>
  <th>A</th>
</tr>
<tr>
  <th>1</th>  
  <td>5</td>
</tr>
<tr>
  <th>2</th>    
  <td>10</td>
</tr>
<tr>
  <th>3</th>  
  <td>3.5</td>
</tr>
<tr>
  <th>4</th>  
  <td>9</td>
</tr>
</table>


Imagine that you write this formula in another cell: `=SUM(A1:A4)`. If you highlight the `A1:A4` part in the formula bar and then you press F9 you'll see the actual values of the range that you selected (In this example, you would see `=SUM({5,10,3.5,9})` . This is super useful for formulas that are extremely large, since you might have an error and this trick allows you to debug each part of the formula to see what's really returning. 

#
#### Shortcut to create new google sheet/doc/slide

If you write  `sheet.new`,  `docs.new` or `slide.new` in your favorite web browser you'll automatically create a new google sheet/doc/slide!

#
#### Shortcut to Find and Replace in Google Sheets

Most frequent shortcut to find things in general is `CTRL + F`, but you can use `CTRL + H` in Google Sheet to use directly the Find and Replace tool.

#
#

### [ES]
# Excel/Google Sheets Formulas

Estas son las formulas y tips que más frecuentemente uso, junto a una descripción de como usarlo. Para aclarar, algunas formulas pueden ser usadas en ambos aplicativos (excel/google sheets) pero puede que requiera algunas modificaciones para que funcionen correctamente.

## Contenido

- [Excel Formulas](#excel-formulas)
  - [Contenido](#content)
      - [Manipulacion de texto](#text-manipulation)
        - [Encontrar si hay un espacio vacio en una celda](#find-if-a-cell-contains-an-empty-space)
        - [Buscar un valor usando dos columnas como referencia](#search-for-a-value-using-two-columns-as-reference)
        - [Eliminar los espacios vacios de las celdas](#remove-empty-spaces-from-a-cell)
        - [Encontrar una palabra específica en una celda y categorizarla.](#search-for-a-keyword-within-a-cell-and-categorize-it-with-a-label)
        - [Reemplazar un #ERROR con algún texto](#replace-an-error-with-some-text)
     - [Formato de Fechas](#date-formats)
        - [Machete para convertir a diferente tipo de formatos](#cheatsheet-to-convert-to-different-type-of-formats) 
     - [Formato de Numeros](#number-formats)
        - [Mostrar los numeros como 1K y 1M en vez de 1,000 y 1,000,000](#show-1k-and-1m-instead-of-1000-and-1000000)  

#

- [Google Sheets Formulas](#google-sheets-formulas)
     - [Array Formulas](#array-formulas)
       - [Auto completar rango con formulas usando una celda como referencia](#fill-range-with-formulas-using-another-cell-as-a-reference)
     - [Importar Data](#data-import)
       - [Importar data filtrada automaticamente de otro google sheet](#import-filtered-data-from-another-google-sheet)    

#

- [Bonus tips y atajos](#bonus-tricks-and-shortcuts)
  - [Chequear directamente el valor de una celda dentro de la barra de formulas](#check-the-value-of-a-cell-inside-the-formula-bar)
  - [Atajo rápido para crear nuevos google sheet/doc/slide](#shortcut-to-create-new-google-sheetdocslide)
  - [Atajo para Buscar y Reemplazar en Google Sheets](#shortcut-to-find-and-replace-in-google-sheets) 
