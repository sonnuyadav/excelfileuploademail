<?php
$realPath = '/home/public_html/';
$month = date('m-Y',strtotime("-1 months"));
$startDate = strtotime(date('01'."-m-Y 00:00:01", strtotime("-1 months")));
$endDate = strtotime(date("t-m-Y 23:59:59", strtotime("-1 months")));

$query = "SELECT count(*) as counters, know_us FROM tbl_enquiry WHERE know_us != '' AND applied_time >=  :startTime AND applied_time <= :endTime GROUP BY know_us ";
$core   = Core::getInstance();
$result = $core->dbh->prepare($query);
$array = array(':startTime'=>$startDate,':endTime'=>$endDate);
$result->execute($array);

$dataRecord = array();
while($record = $result->fetch(PDO::FETCH_ASSOC)){
	array_push($dataRecord, $record);
}

require_once $realPath.'lib/phpExcel/Classes/PHPExcel.php';
require_once $realPath."/mailer/PHPMailerAutoload.php";
require_once $realPath."/mailer/mailsend.php";
$objPHPExcel = new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'Source')
            ->setCellValue('B1', 'Counts');
$objPHPExcel->getActiveSheet()->getStyle('A1:B1')->getFont()->setBold(true);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$counter =2;
foreach($dataRecord as $records){
	$objPHPExcel->setActiveSheetIndex(0)
	            ->setCellValue('A'.$counter, ucfirst($records['know_us']))
            	->setCellValue('B'.$counter, $records['counters']);
            	$counter++;
}
 $filename = 'EnquirySourceCounts_'.$month.'.xlsx';
 $filepath = $realPath.'cron/'.$filename;
 $rd = $objWriter->save($filepath);
 $to  = 'to@gmail.com';  
 sendMail($to,$subject,$body,$cc,$bcc,'',$filepath);