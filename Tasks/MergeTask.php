<?php
namespace ExcelMerge\Tasks;

use ExcelMerge\ExcelMerge;

/**
 * An abstract superclass for all the individual tasks.
 *
 * @package ExcelMerge
 * @property $working_dir
 * @property $result_dir
 */
abstract class MergeTask {
	protected $parent = null;
	protected $sheet_number = null;
	protected $sheet_name = null;

	public function __construct(ExcelMerge $parent) {
		$this->parent = $parent;
	}

	abstract public function merge();

	public function __get($name) {
		switch ($name) {
			case "result_dir":
				return $this->parent->result_dir;
			case "working_dir":
				return $this->parent->working_dir;
		}
		return null;
	}

	public function __set($name, $value) {
		switch ($name) {
			// working_dir and result_dir should be read only, so try to set them on the
			// parent object to throw an error to show that you can't do that.
			case "result_dir":
				$this->parent->result_dir = $value;
				break;
			case "working_dir":
				$this->parent->working_dir = $value;
				break;
		}
	}

	public function set($sheet_number, $sheet_name) {
		$this->sheet_number = $sheet_number;
		$this->sheet_name = $sheet_name;

		return $this; // so we can chain methods, as in $this->set()->merge()
	}
}
