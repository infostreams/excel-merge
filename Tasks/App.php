<?php
namespace ExcelMerge\Tasks;

/**
 * Modifies the "docProps/app.xml" file to contain one more worksheet.
 *
 * @package ExcelMerge\Tasks
 */
class App extends MergeTask {
	public function merge() {
		$filename = "{$this->result_dir}/docProps/app.xml";

		$dom = new \DOMDocument();
		$dom->load($filename);

		/*
		 * 		=> in HeadingPairs/vt:vector/vt:variant[2] set <vt:i4> to {N}
		=> in TitlesOfParts/vt:vector set attribute 'size' to {N}
			=> add
				<vt:lpstr>{New sheet}</vt:lpstr>

		 */

		$xpath = new \DOMXPath($dom);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
		$xpath->registerNamespace("mvt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

		$elems = $xpath->query("//m:HeadingPairs/mvt:vector/mvt:variant[2]/mvt:i4");
		foreach ($elems as $e) {
			$e->nodeValue = $this->sheet_number;
		}

		$elems = $xpath->query("//m:TitlesOfParts/mvt:vector");
		foreach ($elems as $e) {
			$e->setAttribute('size', $this->sheet_number);

			$tag = $dom->createElement('vt:lpstr');
			$tag->nodeValue = $this->sheet_name;

			$e->appendChild($tag);
		}
		$dom->save($filename);
	}
}
