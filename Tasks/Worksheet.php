<?php

namespace ExcelMerge\Tasks;

/**
 * Adds a new worksheet to the merged Excel file
 *
 * @package ExcelMerge\Tasks
 */
class Worksheet extends MergeTask
{
    /**
     * Adds a new worksheet to the merged Excel file
     *
     * @param null $filename The filename of the sheet to copy
     * @param array $shared_strings_mapping
     * @param array $styles_mapping
     *
     * @param array $conditional_styles_mapping
     *
     * @return array
     */
    public function merge($filename = null, $shared_strings_mapping = [], $styles_mapping = [], $conditional_styles_mapping = [])
    {
        if (!file_exists($filename)) {
            return [false, false];
        }
        $new_sheet_number = $this->getSheetCount($this->result_dir) + 1;

        // copy file into place
        $new_name = $this->result_dir . "/xl/worksheets/sheet{$new_sheet_number}.xml";
        if (!is_dir(\dirname($new_name))) {
            mkdir(dirname($new_name));
        }
        copy($filename, $new_name);

        // adjust references to any shared strings
        $sheet = new \DOMDocument();
        $sheet->load($new_name);

        $this->remapSharedStrings($sheet, $shared_strings_mapping);
        $this->remapStyles($sheet, $styles_mapping);
        $this->remapConditionalStyles($sheet, $conditional_styles_mapping);

        $this->disableTabSelected($sheet);

        // save worksheet with adjustments
        $sheet->save($new_name);

        // extract worksheet name
        $sheet_name = $this->extractWorksheetName($filename);

        return [$new_sheet_number, $sheet_name];
    }

    /**
     * @param $dir
     *
     * @return int
     */
    protected function getSheetCount($dir): int
    {
        $existing_sheets = glob("{$dir}/xl/worksheets/sheet*.xml");

        if (\count($existing_sheets) > 0) {
            natsort($existing_sheets);
            $last = basename(end($existing_sheets));

            if (sscanf($last, 'sheet%d.xml', $number)) {
                return $number;
            }
        }

        return 0;
    }

    /**
     * @param $sheet
     * @param $mapping
     */
    protected function remapSharedStrings($sheet, $mapping): void
    {
        $xpath = new \DOMXPath($sheet);
        $xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $shared = $xpath->query("//m:c[@t='s']/m:v");

        if ($shared !== null) {
            foreach ($shared as $tag) {
                $old_id = $tag->nodeValue;

                if (is_numeric($old_id)) {
                    $old_id = (int)$old_id;
                    if (array_key_exists($old_id, $mapping)) {
                        $tag->nodeValue = $mapping[$old_id];
                    }
                }
            }
        }
    }

    /**
     * @param $sheet
     * @param $mapping
     */
    protected function remapStyles($sheet, $mapping): void
    {
        $this->doRemapping($sheet, '//m:c[@s]', 's', $mapping);
        $this->doRemapping($sheet, '//m:row[@s]', 's', $mapping);
        $this->doRemapping($sheet, '//m:col[@style]', 'style', $mapping);
    }

    /**
     * @param $sheet
     * @param $mapping
     */
    protected function remapConditionalStyles($sheet, $mapping): void
    {
        $this->doRemapping($sheet, '//m:conditionalFormatting/m:cfRule[@dxfId]', 'dxfId', $mapping);
    }

    /**
     * @param $sheet
     * @param $xpath_query
     * @param $attribute
     * @param $mapping
     */
    protected function doRemapping($sheet, $xpath_query, $attribute, $mapping): void
    {
        // adjust references to styles
        $xpath = new \DOMXPath($sheet);
        $xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $conditional_styles = $xpath->query($xpath_query);

        if ($conditional_styles !== null) {
            foreach ($conditional_styles as $tag) {
                $old_id = $tag->getAttribute($attribute);
                if (is_numeric($old_id)) {
                    $old_id = (int)$old_id;
                    if (array_key_exists($old_id, $mapping)) {
                        $tag->setAttribute($attribute, $mapping[$old_id]);
                    }
                }
            }
        }
    }

    /**
     * @param $filename
     *
     * @return string
     */
    protected function extractWorksheetName($filename): string
    {
        $workbook = new \DOMDocument();
        $workbook->load(\dirname($filename) . '/../workbook.xml');

        $xpath = new \DOMXPath($workbook);
        $xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        sscanf(basename($filename), 'sheet%d.xml', $number);

        $sheet_name = "Worksheet $number";
        $elems = $xpath->query("//m:sheets/m:sheet[@sheetId='" . $number . "']");
        foreach ($elems as $e) {
            // should be one only
            $sheet_name = $e->getAttribute('name');
            break;
        }

        return $sheet_name;
    }

    private function disableTabSelected($sheet): void
    {
        $xpath = new \DOMXPath($sheet);
        $xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $conditional_styles = $xpath->query('//m:sheetView[@tabSelected]');

        if ($conditional_styles !== null) {
            foreach ($conditional_styles as $tag) {
                $tag->setAttribute('tabSelected', '0');
            }
        }
    }
}
