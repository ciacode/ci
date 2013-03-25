<?php if ( ! defined('BASEPATH')) exit('No direct script access allowed');

class d_excel extends CI_Controller {


	public function index()
	{
		$this->load->view('v_download_excel');
	}
function download_excel(){
	include ("Classes/PHPExcel.php"); 
	//---------------------------------สร้าง object ของ Class  PHPExcel  ขึ้นมาใหม่ 
	$objPHPExcel = new PHPExcel();


	//--------------------------------กำหนดค่าต่างๆ
	$objPHPExcel->getProperties()->setCreator("SIPS");
	$objPHPExcel->getProperties()->setLastModifiedBy("SIPS");
	$objPHPExcel->getProperties()->setTitle("SIPS");
	$objPHPExcel->getProperties()->setSubject("SIPS");
	$objPHPExcel->getProperties()->setDescription("SIPS");

	//--------------------------------Set sharedStyle
	$sharedStyle1 = new PHPExcel_Style();
	$sharedStyle2 = new PHPExcel_Style();

	$sharedStyle1->applyFromArray(
		array('fill' 	=> array(
									'type'		=> PHPExcel_Style_Fill::FILL_SOLID,
									'color'		=> array('argb' => 'FFC0C0C0')
								)
			 ));

	$sharedStyle2->applyFromArray(
		array('fill' 	=> array(
									'type'		=> PHPExcel_Style_Fill::FILL_SOLID,
									'color'		=> array('argb' => 'FFC0C0C0')
								),
			  'borders' => array(
									'bottom'	=> array('style' => PHPExcel_Style_Border::BORDER_THIN),
									'right'		=> array('style' => PHPExcel_Style_Border::BORDER_MEDIUM),
									'top'		=> array('style' => PHPExcel_Style_Border::BORDER_MEDIUM),
									'left'		=> array('style' => PHPExcel_Style_Border::BORDER_MEDIUM)
								)
			 ));


	//---------------------------------เพิ่มข้อมูลเข้าใน Cell
			//----header sheet
			$objPHPExcel->setActiveSheetIndex(0);
			$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:D1');//
			$objPHPExcel->getActiveSheet()->SetCellValue('A1', 'นำเข้าข้อมูลรายกลุ่ม');

			//---header table
			$objPHPExcel->getActiveSheet()->SetCellValue('A3', 'ชื่อ');
			$objPHPExcel->getActiveSheet()->SetCellValue('B3', 'นามสกุล');
			$objPHPExcel->getActiveSheet()->SetCellValue('C3', 'อายุ');
			$objPHPExcel->getActiveSheet()->SetCellValue('D3', 'ที่อยู่');


			//---set style each cell
			$objPHPExcel->getActiveSheet()->setSharedStyle($sharedStyle1, "A1");		
			$objPHPExcel->getActiveSheet()->setSharedStyle($sharedStyle2, "A3:D3");
			$objPHPExcel->getActiveSheet()->getStyle('A1:D3')->getFont()->setBold(true);
			$objPHPExcel->getActiveSheet()->getStyle('A1:D1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objPHPExcel->getActiveSheet()->getStyle('A3:D3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

			//----set width cell
			$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(30);
			$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
			$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(30);
			$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(30);

	//--------------------------------ตั้งชื่อ Sheet
	$objPHPExcel->getActiveSheet()->setTitle('1234');

	//--------------------------------บันทึกไฟล์ Excel 2007
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
	$filename = 'excelExport.xls';
	echo base_url()."upload/excelLoader/".$filename;
	$objWriter->save("upload/excelLoader/".$filename); //ชื่อไฟล์ เมื่อมีการเรียกไฟล์นี้ทำงานก็จะทำการสร้าง ไฟล์ไว้ใน ตำแหน่งของที่กำหนดชื่อไฟล์
	
}
}
?>
