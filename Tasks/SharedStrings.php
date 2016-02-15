<?php
namespace ExcelMerge\Tasks;

/**
 * Consolidates the contents of two 'xl/sharedStrings.xml' files into one, and
 * returns a mapping of how old IDs map onto new IDs.
 *
 * @package ExcelMerge\Tasks
 */
class SharedStrings extends MergeTask {
	public function merge($zip_dir = null) {
		$xml_filename = "/xl/sharedStrings.xml";
		$target_filename = $this->result_dir . $xml_filename;
		$source_filename = $zip_dir . $xml_filename;

		$shared_strings = array();
		$target = new \DOMDocument();
		$target->load($target_filename);
		foreach ($target->documentElement->childNodes as $i=>$ss) {
			// read in current list of shared strings
			$shared_strings[$i] = $ss->nodeValue;
		}

		// add new shared strings, and provide a mapping between old id and new id
		$source = new \DOMDocument();
		$source->load($source_filename);
		$mapping = array();
		foreach ($source->documentElement->childNodes as $i=>$ss) {
			$string = $ss->textContent;

			if (in_array($string, $shared_strings)) {
				$mapping[$i] = array_search($string, $shared_strings);
			} else {
				// we didn't have this string yet
				$shared_strings[] = $string;

				// also add it to $target
				$node = $target->createElement('si');
				$sub = $target->createElement('t');
				$text = $target->createTextNode($string);
				$sub->appendChild($text);

				$node->appendChild($sub);
				$target->documentElement->appendChild($node);

				// record the new mapping
				$mapping[$i] = count($shared_strings) - 1;
			}
		}

		$target->save($target_filename);

		return $mapping;
	}
}