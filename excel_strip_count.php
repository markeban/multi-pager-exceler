<?php

require_once 'excel_reader2.php';
error_reporting(E_ALL ^ E_NOTICE);
$filePaths = [];


		$fileDirectory = "weekly_room_usage/2014/";
		if ($handle = opendir($fileDirectory)) {
			//echo "Directory handle: $handle\n";
			//echo "Entries:<br>";


			while (false !== ($entry = readdir($handle))) {

			array_push($filePaths, $entry);

			}
		    closedir($handle);
}

function cellCount($arg1) {
global $fileDirectory;

	if ($arg1 !== null) {
		$data = new Spreadsheet_Excel_Reader($fileDirectory . $arg1);


		
			$room10 = 0;
			$room20 = 0; 
			$room30 = 0;
			$room40 = 0;
			$room41 = 0;
			$room43 = 0;
			$room50 = 0;
			$room60 = 0;
			$room222 = 0;
			$room226 = 0;
			$room230 = 0;
			$room240 = 0;
			$room244 = 0;
			$room246 = 0;
		
			$rowIndex = 1;
			$num_row = $data->rowcount() + 1;
			while($rowIndex != $num_row){ 
				$colIndex = 1;
				$num_col = $data->colcount() + 1;
					while($colIndex != $num_col){ 
										
					$cell = trim($data->val($rowIndex, $colIndex));
					
					
					
					switch ($cell)
					{
					case "10":
					  $room10 = $room10 + 1;
					  break;
					case "20":
					  $room20 = $room20 + 1;
					  break;
					case "30":
					  $room30 = $room30 + 1;
					  break;
					case "40":
					  $room40 = $room40 + 1;
					  break;
					case "41":
					  $room41 = $room41 + 1;
					  break;
					case "43":
					  $room43 = $room43 + 1;
					  break;  
					case "50":
					  $room50 = $room50 + 1;
					  break;  
					case "60":
					  $room60 = $room60 + 1;
					  break;    
					case "222":
					  $room222 = $room222 + 1;
					  break;  
					case "226":
					  $room226 = $room226 + 1;
					  break;
					case "230":
					  $room230 = $room230 + 1;
					  break;
					case "240":
					  $room240 = $room240 + 1;
					  break;  
					case "244":
					  $room244 = $room244 + 1;
					  break;    
					case "246":
					  $room246 = $room246 + 1;
					  break;
					default:
					}
						
				  $colIndex++;
				  } 
			 $rowIndex++;
			}
			
			$room10 = $room10/12;
			$room20 = $room20/12; 
			$room30 = $room30/12;
			$room40 = $room40/12;
			$room41 = $room41/12;
			$room43 = $room43/12;
			$room50 = $room50/12;
			$room60 = $room60/12;
			$room222 = $room222/6;
			$room226 = $room226/6;
			$room230 = $room230/6;
			$room240 = $room240/6;
			$room244 = $room244/6;
			$room246 = $room246/6;	
			
		$rooms = [$arg1, $room10, $room20, $room30, $room40, $room41, $room43, $room50, $room60, $room222, $room226, $room230, $room240, $room244, $room246];	
		
		return $rooms;
	}	
}
$sheets = [];			
$length = count($filePaths);           	
for ($i =2; $i <= $length; $i++) {

		
	$particularPath = $filePaths[$i];
			
	array_push ($sheets,cellCount($particularPath));
	//print_r ($sheets);
	}

$fileName = (preg_replace('/[^a-zA-Z0-9]/s', '', $fileDirectory) . ".csv") ;	
$file = new SplFileObject($fileName,"w");

foreach (array_filter($sheets) as $fields) {
		$file->fputcsv($fields);
}



echo $fileName . " was created";	

?>
	

