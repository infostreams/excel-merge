<?php
namespace ExcelMerge;

/**
 * Merges two or more Excel files into one.
 *
 * Only Excel 2007 files are supported, so you can only merge .xlsx and .xlxm
 * files. So far, it only seems to work with files that are generated with
 * PHPExcel.
 *
 * @author infostreams https://github.com/infostreams
 *
 * @package ExcelMerge
 * @property $working_dir
 * @property $result_dir
 */
class ExcelMerge {
	protected $files       = array();
	private   $working_dir = null;
	private   $tmp_dir     = null;
	private   $result_dir  = null;
	private   $tasks;
	public    $debug       = false;


	public function __construct($files = array()) {
		// create a temporary directory with an understandable name
		// (comes in use when debugging)
		for ($i=0; $i < 5; $i++) {
			$this->working_dir =
				sys_get_temp_dir() .
				DIRECTORY_SEPARATOR .
				'ExcelMerge-' .
				date('Ymd-His') .
				'-' .
				uniqid() .
				DIRECTORY_SEPARATOR;

			if (!is_dir($this->working_dir)) {
				mkdir($this->working_dir, 0755, true);
				break;
			}
		}

		if (!is_dir($this->working_dir)) {
			trigger_error("Could not create temporary working directory {$this->working_dir}", E_USER_ERROR);
		}


		$this->tmp_dir = $this->working_dir . "tmp" . DIRECTORY_SEPARATOR;
		mkdir($this->tmp_dir, 0755, true);

		$this->result_dir = $this->working_dir . "result" . DIRECTORY_SEPARATOR;
		mkdir($this->result_dir, 0755, true);

		$this->registerMergeTasks();

		foreach ($files as $f) {
			$this->addFile($f);
		}
	}

	public function __destruct() {
		if (!$this->debug) {
			$this->removeTree(realpath($this->working_dir));
		}
	}

	public function addFile($filename) {
		if ($this->isSupportedFile($filename)) {
			if ($this->resultsDirEmpty()) {
				$this->addFirstFile($filename);
			} else {
				$this->mergeWorksheets($filename);
			}
			$this->files[] = $filename;
		}
	}


	/**
	 * Saves the merged file.
	 *
	 * @param null $where
	 * @return string The path and filename to the saved file. The file extension can be
	 * different from the one you provided (!)
	 */
	public function save($where = null) {
		$zipfile = $this->zipContents();
		if ($where === NULL) {
			$where = $zipfile;
		}

		// ignore whatever extension the user might have given us and use the one
		// we obtained in 'zipContents' (i.e. either XLSX or XLSM)
		$where =
			pathinfo($where, PATHINFO_DIRNAME) .
			DIRECTORY_SEPARATOR .
			pathinfo($where, PATHINFO_FILENAME) . "." .
			pathinfo($zipfile, PATHINFO_EXTENSION);

		// move the zipped file to the provided destination
		rename($zipfile, $where);

		// returns the name of the file
		return $where;
	}

	/**
	 * Downloads the merged file
	 *
	 * @param null $download_filename
	 */
	public function download($download_filename = null) {
		$zipfile = $this->zipContents();
		if ($download_filename === NULL) {
			$download_filename = $zipfile;
		}

		// ignore whatever extension the user might have given us and use the one
		// we obtained in 'zipContents' (i.e. either XLSX or XLSM)
		$download_filename =
			pathinfo($download_filename, PATHINFO_FILENAME) . "." .
			pathinfo($zipfile, PATHINFO_EXTENSION);

		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="' . $download_filename . '"');
		header('Cache-Control: max-age=0');
		echo file_get_contents($zipfile);
		unlink($zipfile);
		die;
	}


	protected function addFirstFile($filename) {
		if ($this->resultsDirEmpty()) {
			if ($this->isSupportedFile($filename)) {
				$this->unzip($filename, $this->result_dir);
			}
		} else {
			$this->mergeWorksheets($filename);
		}
	}


	protected function mergeWorksheets($filename) {
		if ($this->resultsDirEmpty()) {
			$this->addFirstFile($filename);
		} else {
			if ($this->isSupportedFile($filename)) {
				$zip_dir = $this->tmp_dir . DIRECTORY_SEPARATOR . basename($filename);
				$this->unzip($filename, $zip_dir);

				$shared_strings = $this->tasks->sharedStrings->merge($zip_dir);
				list($styles, $conditional_styles) = $this->tasks->styles->merge($zip_dir);
				$this->tasks->vba->merge($zip_dir);

				$worksheets = glob("{$zip_dir}/xl/worksheets/sheet*.xml");
				foreach ($worksheets as $s) {
					list($sheet_number, $sheet_name) = $this->tasks->worksheet->merge($s, $shared_strings, $styles, $conditional_styles);

					if ($sheet_number!==false) {
						$this->tasks->workbookRels->set($sheet_number, $sheet_name)->merge();
						$this->tasks->contentTypes->set($sheet_number, $sheet_name)->merge();
						$this->tasks->app->set($sheet_number, $sheet_name)->merge();
						$this->tasks->workbook->set($sheet_number, $sheet_name)->merge();
					}
				}
			}
		}
	}

	protected function registerMergeTasks() {
		$this->tasks = new \stdClass();

		// global tasks
		$this->tasks->sharedStrings = new Tasks\SharedStrings($this);
		$this->tasks->styles = new Tasks\Styles($this);
		$this->tasks->vba = new Tasks\Vba($this);

		// worksheet tasks
		$this->tasks->worksheet = new Tasks\Worksheet($this);
		$this->tasks->workbookRels = new Tasks\WorkbookRels($this);
		$this->tasks->contentTypes = new Tasks\ContentTypes($this);
		$this->tasks->app = new Tasks\App($this);
		$this->tasks->workbook = new Tasks\Workbook($this);
	}


	protected function isSupportedFile($filename, $throw_error = true) {
		$ext = pathinfo($filename, PATHINFO_EXTENSION);
		$is_supported = in_array(strtolower($ext), array('xlsx', 'xlsm'));
		if (!$is_supported && $throw_error) {
			user_error("Can only merge Excel files in .XLSX or .XLSM format. Skipping " . $filename, E_USER_WARNING);
		}

		return $is_supported;
	}

	protected function resultsDirEmpty() {
		return count(array_diff(scandir($this->result_dir), array('.', '..'))) == 0;
	}


	protected function unzip($filename, $directory) {
		$zip = new \ZipArchive();
		$zip->open($filename);
		$zip->extractTo($directory);
		$zip->close();
	}

	protected function removeTree($dir) {
		$result = false;

		$dir = realpath($dir);
		if (strpos($dir, realpath(sys_get_temp_dir())) === 0) {
			$result = true;
			$files = array_diff(scandir($dir), array('.', '..'));
			foreach ($files as $file) {
				if (is_dir("$dir/$file")) {
					$result &= $this->removeTree("$dir/$file");
				} else {
					$result &= unlink("$dir/$file");
				}
			}
			$result &= rmdir($dir);
		}

		return $result;
	}

	protected function zipContents() {
		$zip_directory = realpath($this->result_dir);
		$target_zip = $this->working_dir . DIRECTORY_SEPARATOR . "merged-excel-file";
		$ext = "xlsx";

		$delete = array();

		$zip = new \ZipArchive();
		$zip->open($target_zip, \ZipArchive::CREATE | \ZipArchive::OVERWRITE);

		// Create recursive directory iterator
		/** @var \SplFileInfo[] $files */
		$files = new \RecursiveIteratorIterator(
			new \RecursiveDirectoryIterator($zip_directory),
			\RecursiveIteratorIterator::LEAVES_ONLY
		);

		foreach ($files as $name => $file) {
			// Skip directories (they would be added automatically)
			if (!$file->isDir()) {
				// Get real and relative path for current file
				$filePath = $file->getRealPath();
				if (basename($filePath) != $target_zip) {
					$relativePath = substr($filePath, strlen($zip_directory) + 1);

					// Add current file to archive
					$zip->addFile($filePath, $relativePath);

					$delete[] = $filePath;

					if (basename($filePath) == "vbaProject.bin") {
						// we found VBA code; we change the extension to 'XLSM' to enable macros
						$ext = "xlsm";
					}
				}
			}
		}

		// Zip archive will be created only after closing object
		$zip->close();

		// by default, we delete the files that we put in the zip file
		if (!$this->debug) {
			foreach ($delete as $d) {
				unlink($d);
			}
		}

		// give the zipfile its final name
		rename($target_zip, "$target_zip.$ext");

		return "$target_zip.$ext";
	}

	public function __get($name) {
		switch ($name) {
			case "result_dir":
				return $this->result_dir;
			case "working_dir":
				return $this->working_dir;
		}
		return null;
	}
}