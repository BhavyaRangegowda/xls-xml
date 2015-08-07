<?php

require_once 'Excel/reader.php';

$data = new Spreadsheet_Excel_Reader();
//// Set output Encoding.
$data->setOutputEncoding('CP1251');

//passing the excel to be read 
$data->read('courses.xls');

$main = array();
$sample = array();

//In courses.xls each course info is expaded till 6th row
$j = 6;
$k = 0;


$h = -1;
for($s=0; $s<count($data->sheets)-1; $s++)
{
  $h++;
  $sheetname= $data->boundsheets[$h]['name'];
}
//Loop to read the all the rows in 1st work sheet 
for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++)
{
	//Check the first column is not empty
	if($data->sheets[0]['cells'][$i][1] != "")
	{
		
		//Check if the value is Course Date and do date conversion
		if($data->sheets[0]['cells'][$i][1] == "Course Date")
		{
			
			$date = date("d-m-Y", $data->sheets[0]['cellsInfo'][$i][2]['raw']);
			$spl = explode("-", $date);
			$sample[$data->sheets[0]['cells'][$i][1]] = date("d-M-y", mktime("0", "0", "0", $spl[1], $spl[0]-1, $spl[2]));
		}
		else
		  
			$sample[$data->sheets[0]['cells'][$i][1]] = $data->sheets[0]['cells'][$i][2];
	}
	//if the pointer reaches end of erticular course section(6th row here)
	if($i == $j)
	{
		//assign to main array 
		$main[$k] = $sample;
		//initialise sample arravalues to ""
		$sample = array("Course Type"=>"", "Course Mode"=>"", "Course Name"=>"", "Course Link"=>"", "Course Date"=>"", "Course Duration"=>"");
		
		//Increment j by 7 (6 rows + 1 empty row) to read next perticular set of course info
		$j = $j+7;
		//Increment k
		$k++;
	}
	
}
//Construct xml 
$xml = htmlentities('<onlinecourses>');
$xml.='<br />';
$xml .= htmlentities('<'.$sheetname.'>') . '<br />';
$xml.= htmlentities('<rowdetails><rownames>Course Name,Course Mode,Course Date,Duration</rownames></rowdetails>') . '<br />';
$xml.= htmlentities('<data>');
$xml.='<br />';
for ($k = 0; $k < count($main); $k++)
{
	$course_type = trim($main[$k]["Course Type"]);
		$o_type = $course_type;
	
	$b_type="course";
	
	$course_link = $main[$k]["Course Link"];
	$coursename = htmlentities('<a href='.$course_link.' target="_blank">'.$main[$k]["Course Name"].'</a>');
	$xml .= htmlentities('<'.$b_type.'>') . '<br />';
	
	$xml .= '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'. htmlentities('<column1>'.$coursename.'</column1>') . '<br />';
	$xml .= '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' .htmlentities('<column2>'.htmlentities($main[$k]["Course Mode"]).'</column2>') . '<br />';
	$xml .= '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' .htmlentities('<column3>'.htmlentities($main[$k]["Course Date"]).'</column3>') . '<br />';
	$xml .= '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'. htmlentities('<column4>'.htmlentities($main[$k]["Course Duration"]).'</column4>') . '<br />';
	$xml .= htmlentities('</'.$b_type.'>') . '<br />';
}

$xml .= htmlentities('</data>') . '<br />';
$xml .= htmlentities('</'.$sheetname.'>') . '<br />';
$xml .= htmlentities('</onlinecourses>') . '<br />';

//output the xml
echo $xml;
?>
