<?php
class Util_XLS {


	// property : library for corresponding methods
	private static $libPath = array(
		'array2xls' => array(
			'PhpOffice\PhpSpreadsheet\Spreadsheet',
			'PhpOffice\PhpSpreadsheet\Writer\Xlsx',
		),
		'xls2array' => array(
			__DIR__.'/../../lib/simplexls/0.9.5/src/SimpleXLS.php',
			__DIR__.'/../../lib/simplexlsx/0.8.15/src/SimpleXLSX.php',
		),
	);


	// get (latest) error message
	private static $error;
	public static function error() { return self::$error; }




	/**
	<fusedoc>
		<description>
			export data into excel file (in xlsx format) & save into upload directory
		</description>
		<io>
			<in>
				<!-- parameters -->
				<structure name="$fileData">
					<array name="~worksheetName~">
						<structure name="+" comments="row">
							<string name="~columnName~" />
						</structure>
					</array>
				</structure>
				<string name="$filePath" comments="relative path to upload directory" />
				<structure name="$options">
					<boolean name="showRecordCount" optional="yes" />
					<structure name="columnWidth" optional="yes">
						<array name="~worksheetName~">
							<number name="+" />
						</array>
					</structure>
				</structure>
			</in>
			<out>
				<!-- file output -->
				<file name="~uploadDir~/~filePath~" />
				<!-- return value -->
				<structure name="~return~">
					<string name="path" />
					<string name="url" />
				</structure>
			</out>
		</io>
	</fusedoc>
	*/
	public static function array2xls($fileData, $filePath, $options=[]) {
		// fix swapped parameters
		if ( is_string($fileData) and is_array($filePath) ) list($fileData, $filePath) = array($filePath, $fileData);
		// mark start time
		$startTime = microtime(true);
		// validate library
		foreach ( self::$libPath['array2xls'] as $libClass ) {
			if ( !class_exists($libClass) ) {
				self::$error = '['.__CLASS__.'::'.__FUNCTION__.'] PhpSpreadsheet library is missing ('.$libClass.')<br />Please use <em>composer</em> to install <strong>phpoffice/phpspreadsheet</strong> into your project';
				return false;
			}
		}
		// validate data format
		if ( !is_array($fileData) ) {
			self::$error = '['.__CLASS__.'::'.__FUNCTION__.'] Invalid data structure for Excel (Array is required)';
			return false;
		} elseif ( !empty($fileData) ) {
			$firstWorksheetKey = array_key_first($fileData);
			$firstWorksheetData = $fileData[$firstWorksheetKey];
			if ( !is_array($firstWorksheetData) ) {
				self::$error = '['.__CLASS__.'::'.__FUNCTION__.'] Invalid data structure for Excel (1st level of array is worksheet name, and 2nd level of array is worksheet data)';
				return false;
			}
		}
		// determine output location
		$result = array('path' => self::uploadDir($filePath), 'url'  => self::uploadUrl($filePath));
		if ( $result['path'] === false or $result['url'] === false ) return false;
		// create blank spreadsheet
		$spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
		// go through each worksheet
		$wsIndex = 0;
		foreach ( $fileData as $worksheetName => $worksheet ) {
			// show number of records at worksheet name (when necessary)
			if ( !empty($options['showRecordCount']) and !empty($worksheet) ) {
				$worksheetName .= ' ('.count($worksheet).')';
			}
			// create worksheet
			if ( $wsIndex > 0 ) $spreadsheet->createSheet();
			$spreadsheet->setActiveSheetIndex($wsIndex);
			$activeSheet = $spreadsheet->getActiveSheet();
			$activeSheet->setTitle($worksheetName);
			// all column names (from A to ZZ)
			$alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
			$colNames = str_split($alphabet);
			for ( $i=0; $i<strlen($alphabet); $i++ ) {
				for ( $j=0; $j<strlen($alphabet); $j++ ) {
					$colNames[] = $alphabet[$i].$alphabet[$j];
				}
			}
			// column format
			$activeSheet->getStyle('A:ZZ')->getFont()->setSize(10);
			$activeSheet->getStyle('A:ZZ')->getAlignment()->setWrapText(true);
			$activeSheet->getStyle('A:ZZ')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
			$activeSheet->getStyle('A:ZZ')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
			// header format
			$activeSheet->getStyle('1:1')->getFont()->setBold(true);
			$activeSheet->getStyle('1:1')->getAlignment()->setWrapText(true);
			$activeSheet->getStyle('1:1')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
			$activeSheet->getStyle('1:1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFDDDDDD');
			// column width (when necessary)
			if ( !empty($options['columnWidth'][$worksheetName]) ) {
				foreach ( $options['columnWidth'][$worksheetName] as $key => $val ) {
					$activeSheet->getColumnDimension($colNames[$key])->setWidth($val);
				}
			}
			// output header
			if ( !empty($worksheet) ) {
				$row = $worksheet[0];
				$colIndex = 0;
				foreach ( $row as $key => $val ) {
					$activeSheet->setCellValue($colNames[$colIndex].'1', $key);
					$colIndex++;
				}
			}
			// output each row of data
			foreach ( $worksheet as $rowIndex => $row ) {
				$rowNumber = $rowIndex + 2;
				// go through each column
				$colIndex = 0;
				foreach ( $row as $key => $val ) {
					$activeSheet->setCellValue($colNames[$colIndex].$rowNumber, $val);
					$colIndex++;
				} // foreach-col
			} // foreach-row
			$wsIndex++;
			// focus first cell (when finished)
			$activeSheet->getStyle('A1');
		} // foreach-worksheet
		// mark end time
		$endTime = microtime(true);
		$et = round(($endTime-$startTime)*1000);
		// show execution time at last worksheet
		$spreadsheet->createSheet();
		$spreadsheet->setActiveSheetIndex( count($fileData) );
		$activeSheet = $spreadsheet->getActiveSheet();
		$activeSheet->setTitle('et ('.$et.'ms)');
		// focus first worksheet (when finished)
		$spreadsheet->setActiveSheetIndex(0);
		// write to report
		$writer = new PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
		$writer->save($result['path']);
		// done!
		return $result;
	}


} // class