<?

//error_reporting(0);

// $workdir = "d:\\Visual Studio 2008\\Projects\\balsverka\\Debug";


$workdir = str_replace("\\", "/" , $work_dir);

$balbars_file = str_replace("\\", "/" , $balbars_file);
$autobars_file = str_replace("\\", "/" , $autobars_file);
$balkazna_file = str_replace("\\", "/" , $balkazna_file);

$file = $workdir."/blank/balsverka_XXXX.xls";

$newfile = $workdir."/balsverka.xls";


if (!copy($file, $newfile)) {
    echo "failed to copy $file...\n";
}


$balbars_file_lines = file($balbars_file);
$autobars_file_lines = file($autobars_file);
$balkazna_file_data = dbase_open($balkazna_file, 0);

$sheet1 = "accounts";

$excel_app = new COM("Excel.application") or Die ("Did not connect");

$excel_app->Visible = 1;

$Workbook = $excel_app->Workbooks->Open($newfile) or Die("Did not open $filename $Workbook");
$Worksheet = $Workbook->Worksheets($sheet1);
$Worksheet->activate;


$i=2;

$excel_result_balacc = '0000';
while ($excel_result_balacc !='')
{

$coord_balacc = "A".$i;
$coord_balbars = "B".$i;
$coord_autobars = "C".$i;
$coord_balkazna = "D".$i;


$excel_cell_balacc = $Worksheet->Range($coord_balacc);
$excel_cell_balacc->activate;
$excel_result_balacc = $excel_cell_balacc->value;

foreach ($balbars_file_lines as $balbars_file_line_num)
{

$balbars_file_line_num_data = explode("  ", trim($balbars_file_line_num));

if ( trim($balbars_file_line_num_data[0]) == trim($excel_result_balacc))

{


$excel_cell_balbars = $Worksheet->Range($coord_balbars);

$excel_cell_balbars->activate;

$balbars_data_tmp = explode(" ", trim($balbars_file_line_num_data[3]));
$balbars_data = implode($balbars_data_tmp);

$excel_cell_balbars->value = $balbars_data;
//$excel_cell_balbars->value = "gg";

//echo $pieces[1].":".$pieces[2]."\n"; // piece1

}

}

foreach ($autobars_file_lines as $autobars_file_line_num)
{

$autobars_file_line_num_data = explode("                ", trim($autobars_file_line_num));

if ( trim($autobars_file_line_num_data[0]) == trim($excel_result_balacc))

{

$excel_cell_autobars = $Worksheet->Range($coord_autobars);
$excel_cell_autobars->activate;

$autobars_data_tmp = explode(" ", trim($autobars_file_line_num_data[2]));
$autobars_data = implode($autobars_data_tmp);


$excel_cell_autobars->value = $autobars_data;

}


}


$balkazna_file_data_record_num = dbase_numrecords($balkazna_file_data);

  for ($y = 1; $y <= $balkazna_file_data_record_num; $y++)
  {

   $row = dbase_get_record_with_names($balkazna_file_data, $y);

  if ($row['BALANCE'] == trim($excel_result_balacc))
          {
          $excel_cell_balkazna = $Worksheet->Range($coord_balkazna);
          $excel_cell_balkazna->activate; $excel_cell_balkazna->value = $row['DB_DAY'];
          }

    }


if ($excel_result_balacc ==''){break;}
//echo $excel_result_balacc."\n";



$i = $i + 1;

}

dbase_close($balkazna_file_data);


// closing excel

$excel_app->ActiveWorkbook->Save();

$excel_app->Quit();

// free the object
//$excel_app->Release();

$excel_app = null;

?>
