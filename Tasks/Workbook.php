<?php
namespace ExcelMerge\Tasks;

/**
 * Modifies the "xl/workbook.xml" file to contain one more worksheet.
 *
 * @package ExcelMerge\Tasks
 */
class Workbook extends MergeTask {
	public function merge() {
		/**
		 * 	7. xl/workbook.xml
		 *         => add
		 *            <sheet name="{New sheet}" sheetId="{N}" r:id="rId{N}"/>
		 */
		$filename = "{$this->result_dir}/xl/workbook.xml";
		$dom = new \DOMDocument();
		$dom->load($filename);

		$xpath = new \DOMXPath($dom);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		$elems = $xpath->query("//m:sheets");
		foreach ($elems as $e) {
			$tag = $dom->createElement('sheet');
			$tag->setAttribute('name', $this->sheet_name);
			$tag->setAttribute('sheetId', $this->sheet_number);
			$tag->setAttribute('r:id', "rId" . $this->sheet_number);

			$e->appendChild($tag);
			break;
		}

		// make sure all worksheets have the correct rId - we might have assigned them new ids
		// in the Tasks\WorkbookRels::merge() method
		$elems = $xpath->query("//m:sheets/m:sheet");
		foreach ($elems as $e) {
			$e->setAttribute("r:id", "rId" . ($e->getAttribute("sheetId")));
		}

		$dom->save($filename);
	}

}
