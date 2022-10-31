<?php
class Util_Excel {


	// property : library for corresponding methods
	private static $libPath = array(
		'array2xls' => '',
		'xls2array' => ''
	);


	// get (latest) error message
	private static $error;
	public static function error() { return self::$error; }


} // class