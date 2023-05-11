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
				<string name="$filePath" optional="yes" comments="relative path to upload directory; download directly when not specified" />
				<structure name="$options">
					<boolean name="multipleWorksheets" optional="yes" default="false" />
					<boolean name="showRecordCount" optional="yes" default="false" />
					<structure name="columnWidth" optional="yes" default="~emptyArray~">
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
	public static function array2xls($fileData, $filePath=null, $options=[]) {
		// mark start time
		$startTime = microtime(true);
		// fix swapped parameters
		if ( isset($filePath) and is_string($fileData) and is_array($filePath) ) list($fileData, $filePath) = array($filePath, $fileData);
		// default options
		$options['multipleWorksheets'] = $options['multipleWorksheets'] ?? false;
		$options['showRecordCount'] = $options['showRecordCount'] ?? false;
		$options['columnWidth'] = $options['columnWidth'] ?? [];
		// wrap by an extra layer of array (when single worksheet)
		if ( !$options['multipleWorksheets'] ) {
			$fileData = array('Untitled' => $fileData);
			$options['columnWidth'] = array('Untitled' => $options['columnWidth']);
		}
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
		// create blank spreadsheet
		$spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
		// go through each worksheet
		$wsIndex = 0;
		foreach ( $fileData as $worksheetName => $worksheet ) {
			// show number of records at worksheet name (when necessary)
			if ( $options['showRecordCount'] and !empty($worksheet) ) {
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
				$row = $worksheet[array_key_first($worksheet)];
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
		// determine output location
		// ===> when file path not specified
		// ===> output to temp file
		if ( $filePath ) {
			$result = array('path' => Util::uploadDir($filePath), 'url' => Util::uploadUrl($filePath));
			if ( $result['path'] === false or $result['url'] === false ) {
				self::$error = '['.__CLASS__.'::'.__FUNCTION__.'] '.Util::error();
				return false;
			}
		} else {
			$uuid = Util::uuid();
			if ( $uuid === false ) {
				self::$error = '['.__CLASS__.'::'.__FUNCTION__.'] '.Util::error();
				return false;
			}
			$result = array('path' => $uuid.'.xls', 'url' => $uuid.'.xls');
		}
		// write to report
		$writer = new PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
		$writer->save($result['path']);
		// when file path not specified
		// ===> download directly
		if ( !$filePath ) {
			$streamed = Util::streamFile($result['path'], [ 'download' => true, 'deleteAfterward' => true ]);
			if ( $streamed === false ) {
				self::$error = '['.__CLASS__.'::'.__FUNCTION__.'] '.Util::error();
				return false;
			}
		}
		// done!
		return $result;
	}




	/**
	<fusedoc>
		<description>
			convert csv/xls/xlsx to array
			===> use first row as column name (when necessary)
			===> use snake-case for column name (e.g. this_is_col_name)
		</description>
		<io>
			<in>
				<path name="$file" comments="excel file path" />
				<structure name="$options">
					<number name="worksheet" default="0" comments="starts from zero" />
					<number name="startRow" default="1" comments="starts from one" />
					<boolean name="firstRowAsHeader" default="true" />
					<boolean name="convertHeaderCase" default="true" />
				</structure>
			</in>
			<out>
				<array name="~return~">
					<structure name="+">
						<string name="~columnName~" />
					</structure>
				</array>
			</out>
		</io>
	</fusedoc>
	*/
	public static function xls2array($file, $options=[]) {
		// default options
		$options['startRow'] = $options['startRow'] ?? 1;
		$options['worksheet'] = $options['worksheet'] ?? 0;
		$options['firstRowAsHeader'] = $options['firstRowAsHeader'] ?? true;
		$options['convertHeaderCase'] = $options['convertHeaderCase'] ?? true;
		// load library
		foreach ( self::$libPath['xls2array'] as $path ) {
			if ( !is_file($path) ) {
				self::$error = '['.__CLASS__.'::'.__FUNCTION__.'] SimpleXLSX library is missing ('.$path.')';
				return false;
			}
			require_once($path);
		}
		// validation
		$fileExt = strtoupper( pathinfo($file, PATHINFO_EXTENSION) );
		if ( !is_file($file) ) {
			self::$error = '['.__CLASS__.'::'.__FUNCTION__.'] File not found ('.$file.')';
			return false;
		} elseif ( !in_array($fileExt, ['XLSX','XLS','CSV']) ) {
			self::$error = '['.__CLASS__.'::'.__FUNCTION__.'] File type <strong><em>'.$fileExt.'</em></strong> is not supported';
			return false;
		}
		// parse csv by php
		if ( $fileExt == 'CSV' ) {
			$data = file_get_contents($file);
			$data = mb_convert_encoding($data, 'UTF-8', mb_detect_encoding($data, 'UTF-8, ISO-8859-1', true));
			$data = array_map('str_getcsv', explode(PHP_EOL, $data));
		// parse excel by library
		} else {
			$data = call_user_func('Simple'.$fileExt.'::parse', $file);
			if ( $data === false ) {
				self::$error = '['.__CLASS__.'::'.__FUNCTION__.'] '.call_user_func('Simple'.$fileExt.'::parseError');
				return false;
			}
		}
		// extract data from specific worksheet (when necessary)
		if ( method_exists($data, 'rows') ) $data = $data->rows($options['worksheet']);
		for ( $i=0; $i<($options['startRow']-1); $i++ ) if ( isset($data[$i]) ) unset($data[$i]);
		$data = array_values($data);
		// validation
		// ===> simply return when no data
		// ===> simply return when no need to apply first row as header
		if ( empty($data) or !$options['firstRowAsHeader'] ) return $data;
		// get column name from first row
		$colNames = $data[0];
		unset($data[0]);
		$data = array_values($data);
		// convert column name into snake case
		if ( $options['convertHeaderCase'] ) {
			$colNames = array_map('strtolower', $colNames);
			foreach ( $colNames as $i => $val ) {
				$val = strtolower($val);
				$val = preg_replace( '/[^a-z0-9]/i', ' ', $val);
				$val = preg_replace('!\s+!', ' ', $val);
				$val = str_replace(' ', '_', $val);
				$val = trim($val, '_');
				$colNames[$i] = $val;
			}
		}
		// go through each row and create new record
		$result = array();
		foreach ( $data as $row => $rowData ) {
			$item = array();
			foreach ( $colNames as $colIndex => $colName ) {
				$item[$colName] = isset($rowData[$colIndex]) ? $rowData[$colIndex] : '';
			}
			$result[] = $item;
		}
		// clean-up data
		foreach ( $result as $row => $rowData ) {
			foreach ( $rowData as $col => $val ) {
				$result[$row][$col] = trim($val);
			}
		}
		// done!
		return $result;
	}


} // class