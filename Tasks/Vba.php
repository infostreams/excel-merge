<?php
namespace ExcelMerge\Tasks;

/**
 * Adds a workbook's VBA code to the merged Excel file.
 *
 * If more than one of the source files contain VBA code, then the merged
 * file will contain only the code from the source file that was merged
 * last.
 *
 * @package ExcelMerge\Tasks
 */
class Vba extends MergeTask {

	public function merge($zip_dir = null) {
		if ($this->insertFile($zip_dir)) {
			// successfully copied VBA code into merged file
			$this->addWorkbookRelation();
			$this->registerContentType();
		}
	}

	protected function insertFile($zip_dir) {
		$filename = "/xl/vbaProject.bin";
		$target_filename = $this->result_dir . $filename;
		$source_filename = $zip_dir . $filename;

		if (file_exists($source_filename)) {
			if (file_exists($target_filename)) {
				// if the target file already exists, try to delete it first
				@unlink($target_filename);
			}
			if (!file_exists($target_filename)) {
				// we only try to copy the file to the target location
				// if there's no identically named file there already
				if (copy($source_filename, $target_filename)) {
					return true;
				}
			}
		}

		return false;
	}

	protected function addWorkbookRelation() {
		// Add (if necessary) the following to _rels/workbook.xml.rels:
		//  <Relationship Id="rId{N}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
		$rels_file = $this->result_dir . "xl/_rels/workbook.xml.rels";

		$doc = new \DOMDocument();
		$doc->load($rels_file);

		$xpath = new \DOMXPath($doc);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/package/2006/relationships");

		$elems = $xpath->query("//m:Relationship[@Target='vbaProject.bin']");

		if ($elems->length == 0) {
			$ids = $xpath->query("//m:Relationship");

			$node = $doc->createElement("Relationship");
			$node->setAttribute("Id", "rId" . ($ids->length + 1));
			$node->setAttribute("Type", "http://schemas.microsoft.com/office/2006/relationships/vbaProject");
			$node->setAttribute("Target", "vbaProject.bin");
			$doc->documentElement->appendChild($node);

			$doc->save($rels_file);
		}
	}

	protected function registerContentType() {
		// and add (if necessary) the following to [Content_Types].xml:
		// <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
		$content_types_file = $this->result_dir . "[Content_Types].xml";

		$doc = new \DOMDocument();
		$doc->load($content_types_file);

		$xpath = new \DOMXPath($doc);
		$xpath->registerNamespace("m", "http://schemas.openxmlformats.org/package/2006/content-types");

		$elems = $xpath->query("//m:Default[@Extension='bin']");
		if ($elems->length == 0) {
			$node = $doc->createElement("Default");
			$node->setAttribute("Extension", "bin");
			$node->setAttribute("ContentType", "application/vnd.ms-office.vbaProject");

			$doc->documentElement->appendChild($node);

			$doc->save($content_types_file);
		}
	}
}