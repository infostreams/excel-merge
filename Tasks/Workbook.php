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
		
		// Caroline Clep: this is breaking the result file - need to make sure we don't touch the sheets ids and only update the external links
		//$elems = $xpath->query("//m:sheets/m:sheet");
		//foreach ($elems as $e) {
		//	$e->setAttribute("r:id", "rId" . ($e->getAttribute("sheetId")));
		//}
		
		$relfilename = "{$this->result_dir}/xl/_rels/workbook.xml.rels";
		$reldom = new \DOMDocument();
		$reldom->load($relfilename);

		$relxpath = new \DOMXPath($reldom);
		$relxpath->registerNamespace("m", "http://schemas.openxmlformats.org/package/2006/relationships");
		$relelems = $relxpath->query("//m:Relationship");


		$elems = $xpath->query("//m:externalReference");
		$refId = 1;
		foreach ($elems as $e)
		{
			foreach ($relelems as $rele)
			{
				if ($rele->getAttribute("Target") === "externalLinks/externalLink" . $refId . ".xml")
				{
					$e->setAttribute("r:id", $rele->getAttribute("Id"));
					break;
				}
			}
			$refId++;
		}
		// Caroline Clep: End of fix

		$dom->save($filename);
	}

}
