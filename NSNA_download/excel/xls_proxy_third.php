<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Asia/Seoul');
if (PHP_SAPI == 'cli')
	die('This example should only be run from a Web Browser');

/** Include PHPExcel */
require_once '../Classes/PHPExcel.php';

// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');
// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0
header('Content-Disposition: attachment;filename="article.xls"');

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
// Set document properties
$objPHPExcel->getProperties()->setCreator("NSNA")
							 ->setLastModifiedBy("NSNA")
							 ->setTitle("NewsSource XLS")
							 ->setSubject("NewsSource XLS")
							 ->setDescription("NewsSource XLS")
							 ->setKeywords("NewsSource XLS")
							 ->setCategory("NewsSource XLS");

function trim_str($val) {
	$ret = (string) $val;
	$ret = trim($ret);
	return $ret;
}

//$type = $_GET['type'];
$keyword = urlencode($_GET['keyword']);// 검색 키워드를 의미. 주의사항: 다중키워드 검색을 할 경우 키워드 토큰을 white space가 아닌 +기호로 바꿔주세요. 예) 박근혜+대통령
$begin = $_GET['begin']; //검색 시작점. 양식 YYYYMMDD
$end = $_GET['end'];
$providers= "";
if (isset($_GET['providers'])) 
	$providers = urlencode($_GET['providers']); //언론사 필터링

$begin_date = DateTime::createFromFormat('Ymd', $begin);
$end_date = DateTime::createFromFormat('Ymd', $end);
$unit = $end_date->diff($begin_date)->days;
$url = "http://147.47.125.161:9999/NSNA_FrontEnd/JSP/DownloadExcel.jsp?keyword=".$keyword."&begin=".$begin."&end=".$end."&type=third&providers=".$providers."&unit=".$unit;

$xml = simplexml_load_file($url);
//$xml=simplexml_load_file("example.xml");

foreach($xml -> UNIT  as $unit_child) {
	$active_sheet = $objPHPExcel->setActiveSheetIndex(0)->setTitle('Article');
	$row_num = 2;
	$active_sheet->setCellValue('A1', 'ART_ID')
                                ->setCellValue('B1', 'ART_CATEGORY')
				->setCellValue('C1', 'ART_HEADLINE')
				->setCellValue('D1', 'ART_DATE')
				->setCellValue('E1', 'ART_PROVIDER')
				->setCellValue('F1', 'ART_CONTENT');

	foreach($unit_child-> ARTICLE -> children() as $child) {
		$child_name = $child->getName();
		if ($child_name == "ART") {
			$active_sheet->setCellValueByColumnAndRow(0, $row_num, trim_str($child->ART_ID));
                        $active_sheet->setCellValueByColumnAndRow(1, $row_num, trim_str($child->ART_CATEGORY));
			$active_sheet->setCellValueByColumnAndRow(2, $row_num, trim_str($child->ART_HEADLINE));
			$active_sheet->setCellValueByColumnAndRow(3, $row_num, trim_str($child->ART_DATE));
			$active_sheet->setCellValueByColumnAndRow(4, $row_num, trim_str($child->ART_PROVIDER));
			$contents = "";
			foreach($child -> ART_CONTENT -> children() as $content) {
				$contents.= trim_str($content);
			}
			$active_sheet->setCellValueByColumnAndRow(5, $row_num, trim_str($contents));
			$row_num += 1;
		}
	}
	break;
}

$active_sheet = $objPHPExcel->createSheet()->setTitle('Query_Info');
$row_num = 2;
$active_sheet->setCellValue('A1', 'QUERY')
			->setCellValue('B1', 'BEGIN')
			->setCellValue('C1', 'END')
			->setCellValue('D1', 'PROVIDERS')
			->setCellValue('A2', urldecode($keyword))
			->setCellValue('B2', $begin)
			->setCellValue('C2', $end)
			->setCellValue('D2', urldecode($providers));
// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;



?>
