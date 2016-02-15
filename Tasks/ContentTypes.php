<?php
namespace ExcelMerge\Tasks;

/**
 * Modifies the "[Content_Types].xml" file to contain one more worksheet.
 *
 * @package ExcelMerge\Tasks
 */
class ContentTypes extends MergeTask {
	public function merge() {
		$filename = "{$this->result_dir}/[Content_Types].xml";

		$dom = new \DOMDocument();
		$dom->load($filename);

		$tag = $dom->createElement("Override");
		$tag->setAttribute('PartName', "/xl/worksheets/sheet{$this->sheet_number}.xml");
		$tag->setAttribute('ContentType', "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");

		$dom->documentElement->appendChild($tag);

		$dom->save($filename);
	}
}
