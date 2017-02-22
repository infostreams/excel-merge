<?php
namespace ExcelMerge\Tasks;

/**
 * Modifies the "xl/_rels/workbook.xml.rels" file to contain one more worksheet.
 *
 * @package ExcelMerge\Tasks
 */
class WorkbookRels extends MergeTask {
	public function merge() {
		/**
		 *  xl/_rels/workbook.xml.rels
		 *  => in 'Relationships'
		 *  => add
		 *  <Relationship Id="rId{N}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{N}.xml"/>
                 *
	      	 *  => Renumber all rId{X} values to rId{X+1} where X >= N
		 *
		 * -> Re-order and re-number so that we first list all the sheets, and then the rest
		 */

		$filename = "{$this->result_dir}/xl/_rels/workbook.xml.rels";
		$dom = new \DOMDocument();
		$dom->load($filename);

		$xpath = new \DOMXPath($dom);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/package/2006/relationships");
		$elems = $xpath->query("//m:Relationship");

		$rest_id = $this->sheet_number + 1;
		foreach ($elems as $e) {
			$type = $e->getAttribute("Type");
			$is_worksheet = (strpos($type, "worksheet")!==false);

			if ($is_worksheet) {
				sscanf($e->getAttribute("Target"), "worksheets/sheet%d.xml", $sheet_nr);
				$e->setAttribute("Id", "rId" . ($sheet_nr));
			} else {
				$e->setAttribute("Id", "rId" . ($rest_id++));
			}
		}

		$new_rid = "rId" . $this->sheet_number;
		$tag = $dom->createElement('Relationship');
		$tag->setAttribute('Id', $new_rid);
		$tag->setAttribute('Type', "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
		$tag->setAttribute('Target', "worksheets/sheet" . $this->sheet_number . ".xml");

		$dom->documentElement->appendChild($tag);

		$dom->save($filename);
	}
}
