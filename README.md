VBA-Implode
===========
<h2>Origination</h2>
This function was originally intended to implode data from multiple rows matching one id into one row. 
I.e., from
<table><thead><tr><th scope="col">ID</th>
<th scope="col">Picture</th>
</tr></thead><tbody><tr><td>a</td>
<td>a-1.jpg</td>
</tr><tr><td>a</td>
<td>a-2.jpg</td>
</tr><tr><td>a</td>
<td>a-3.jpg</td>
</tr><tr><td>b</td>
<td>b-1.jpg</td>
</tr><tr><td>b</td>
<td>b-2.jpg</td>
</tr><tr><td>c</td>
<td>c-1.jpg</td>
</tr><tr><td>c</td>
<td>c-2.jpg</td>
</tr><tr><td>c</td>
<td>c-3.jpg</td>
</tr></tbody></table>

to

<table><thead><tr><th scope="col">ID</th>
<th scope="col">Picture</th>
</tr></thead><tbody><tr><td>a</td>
<td>dir/a-1.jpg | dir/a-2.jpg | dir/a-3.jpg</td>
</tr><tr><td>b</td>
<td>dir/b-1.jpg | dir/b-2.jpg</td>
</tr><tr><td>c</td>
<td>dir/c-1.jpg | dir/c-2.jpg | dir/c-3.jpg</td>
</tr></tbody></table>

The reasoning behind this was to make it easy to import into a Drupal site using Feeds. 

<h2>Use</h2>
This VB script is intended for dropping into a macro-enabled Microsoft Excel sheet (or any other spreadsheet program that reads visual basic macros). (To read how to start a macro in Excel, click <a href="http://www.wikihow.com/Write-a-Simple-Macro-in-Microsoft-Excel">here</a>.) 

<b>This is the original code that I used. You will have to adjust the code to make it fit your situation. This will be modified later so to make it easier.</b>

<ol>
<li>Change "A2" in <pre>Call IMPLODE_FXN(Sheet1.Range("A2"))</pre> to whatever cell your column of IDs start. </li>
<li>If you have over 3,000 records, change 5000 in <pre>For i = 0 to 5000</pre> to something around 10,000. <ul><li>I am not sure why it won't execute properly when you put in the number of rows you have, and when I tried to put in a function to call in the number of rows in the sheet, I got an error.</li></ul></li>
<li>Change the value of what is being written to fit your desired file path.</li>
<li>Run!</li>
</ol>

If you have any questions about modifying any of the code, or have questions as to why something was done, please feel free to email me at <a href="mailto:logan@loganfarr.com">logan@loganfarr.com</a>. I will try to answer as many of your questions as best as I can.

Regards,
<a href="http://loganfarr.com">Logan</a>
