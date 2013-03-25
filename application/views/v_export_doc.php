<?php
//word
header("Content-Type: application/msword");
header('Content-disposition: inline; filename="exportWord.doc"');

//excel
//header('Content-Type: application/xls');
//header('Content-Disposition: attachment; filename="exportExcel.xls"');
?>
<html>
<body>
<Table border=2 align="center" >
<tr >
<th >ชื่่อ</th>
<th>นาสกุล</th>
</tr>
<tr>
<td>ทรงพล</td>
<td>รัตนานนท์</td>
</tr>
<tr >
<td>โนบิ</td>
<td>โนบิตะ</td>
</tr>
</table>
</body>
</html>