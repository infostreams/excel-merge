<?php

namespace ExcelMerge\Tasks;

/**
 * Consolidates the contents of two 'xl/styles.xml' files into one, and
 * returns two mappings:
 *
 * 1. a mapping of how old style IDs map onto new style IDs
 * 2. a mapping of how old 'conditional style' IDs map onto new 'conditional style' IDs
 *
 * @package ExcelMerge\Tasks
 */
class Styles extends MergeTask
{
    protected $style_tags = ['fonts', 'fills', 'borders', 'dxfs'];
    private $definedStyles = [];
    private $newFmtId = 1000;

    /**
     * @param null $zip_dir
     *
     * @return array
     */
    public function merge($zip_dir = null)
    {
        $xml_filename = '/xl/styles.xml';
        $existing_filename = $this->result_dir . $xml_filename;
        $source_filename = $zip_dir . $xml_filename;

        // get hash signature for each entry in 'numfmt', 'fonts', 'fills' and 'borders'
        // see if there are any new ones
        // - if so, add them and store the id. Make sure to update the 'count' attribute in the parent tag
        // - if it already existed, get the id
        $existing_dom = new \DOMDocument();
        $existing_dom->load($existing_filename);

        $existing_xpath = new \DOMXPath($existing_dom);
        $existing_xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        $styles = $this->getStyles($existing_xpath);

        $source_dom = new \DOMDocument();
        $source_dom->load($source_filename);

        // re-assign xpath to work on source doc
        $source_xpath = new \DOMXPath($source_dom);
        $source_xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // iterate all the style tags in document that we want to merge in
        [$mapping, $styles] = $this->addNewStyles($source_xpath, $styles);

        // replace styles from existing styles.xml document with the merged styles
        $this->replaceStyleTags($styles, $existing_xpath);

        // now go through the 'cellXfs' tags. Update the references to 'fontId', 'numFmtId',
        // 'fillId', and 'borderId'. Generate a tag for each style that we're importing.
        //
        // If it already existed, note the id. If it didn't exist, add it and store the id.
        // Return this mapping of ids
        [$defined_styles, $styles_mapping] = $this->rewriteCells($existing_xpath, $source_xpath, $mapping);

        // write the new styles list
        $this->replaceStylesList($defined_styles, $existing_xpath);

        // save the merged style file
        $existing_dom->save($existing_filename);

        // return a mapping of how style ids in this workbook relate to style ids in the merged workbook
        return [$styles_mapping, $mapping['dxfs']];
    }

    /**
     * @param $existing_xpath
     *
     * @return array
     */
    protected function getStyles($existing_xpath): array
    {
        $existing_styles = [];
        $tags = array_merge($this->style_tags, ['cellStyles', 'numFmts']);
        foreach ($tags as $tag) {
            $elems = $existing_xpath->query("//m:{$tag}");
            $existing_styles[$tag] = [];
            if ($elems->length > 0) {
                if ($elems->item(0)->hasChildNodes()) {
                    foreach ($elems->item(0)->childNodes as $id => $style) {
                        $existing_styles[$tag][$id] = [
                            'node' => $style,
                            'string' => $style->C14N(true, false),
                            'id' => $id,
                        ];
                    }
                }
            }
        }
        return $existing_styles;
    }

    /**
     * @param \DOMXPath $source_xpath The document to add styles from
     * @param $existing_styles
     *
     * @return array
     */
    protected function addNewStyles($source_xpath, $existing_styles): array
    {
        $mapping = [];
        foreach ($this->style_tags as $tag) {
            $elems = $source_xpath->query("//m:{$tag}");
            $mapping[$tag] = [];
            if ($elems && $elems->item(0) && $elems->item(0)->hasChildNodes()) {
                foreach ($elems->item(0)->childNodes as $id => $style) {
                    $string = $style->C14N(true, false);
                    foreach ($existing_styles[$tag] as $e) {
                        if ($e['string'] === $string) {
                            // this is an existing style
                            $mapping[$tag][$id] = $e['id'];
                            continue 2; // continue to next style
                        }
                    }

                    // this is a new style
                    $new_id = \count($existing_styles[$tag]);

                    $existing_styles[$tag][$new_id] = [
                        'node' => $style,
                        'string' => $style->C14N(true, false),
                        'id' => $new_id,
                    ];
                    $mapping[$tag][$id] = $new_id;
                }
            }
        }

        $tag = 'cellStyles';
        $elems = $source_xpath->query("//m:{$tag}");
        $mapping[$tag] = [];
        if ($elems && $elems->item(0) && $elems->item(0)->hasChildNodes()) {
            foreach ($elems->item(0)->childNodes as $id => $style) {
                foreach ($existing_styles[$tag] as $e) {
                    if ($style->getAttribute('name') === $e['node']->getAttribute('name')) {
                        // this is an existing style
                        $mapping[$tag][$id] = $e['id'];
                        continue 2; // continue to next style
                    }
                }

                // this is a new style
                $new_id = \count($existing_styles[$tag]);

                $existing_styles[$tag][$new_id] = [
                    'node' => $style,
                    'string' => $style->C14N(true, false),
                    'id' => $new_id,
                ];
                $mapping[$tag][$id] = $new_id;
            }
        }

        $tag = 'numFmts';
        $elems = $source_xpath->query("//m:{$tag}");
        $mapping[$tag] = [];
        if ($elems && $elems->item(0) && $elems->item(0)->hasChildNodes()) {
            foreach ($elems->item(0)->childNodes as $id => $style) {
                $mergedFmtId = (int)$style->getAttribute('numFmtId');
                foreach ($existing_styles[$tag] as $e) {
                    if ($style->getAttribute('formatCode') === $e['node']->getAttribute('formatCode')) {
                        // this is an existing style
                        $originFmtId = (int)$e['node']->getAttribute('numFmtId');
                        $mapping[$tag][$mergedFmtId] = $originFmtId;
                        continue 2; // continue to next style
                    }
                }

                // this is a new style
                $this->newFmtId++;
                $new_id = \count($existing_styles[$tag]);
                $newStyle = $style;
                $newStyle->setAttribute('numFmtId', $this->newFmtId);
                $existing_styles[$tag][$new_id] = [
                    'node' => $newStyle,
                    'string' => $newStyle->C14N(true, false),
                    'id' => $new_id,
                ];
                $mapping[$tag][$mergedFmtId] = $this->newFmtId;
            }
        }
        return [$mapping, $existing_styles];
    }

    /**
     * @param $existing_styles
     * @param \DOMXPath $xpath
     */
    protected function replaceStyleTags($existing_styles, $xpath): void
    {
        foreach ($existing_styles as $tag => $styles) {
            $elems = $xpath->query("//m:{$tag}");

            if ($elems->length > 0) {
                $elem = $elems->item(0);
                while ($elem->hasChildNodes()) {
                    $elem->removeChild($elem->firstChild);
                }
                foreach ($styles as $s) {
                    $elem->appendChild($xpath->document->importNode($s['node'], true));
                }
                $elem->setAttribute('count', \count($styles));
            }
        }
    }

    /**
     * @param \DOMXPath $dompath
     *
     * @return array
     */
    private function getFileStyles(\DOMXPath $dompath): array
    {
        $styles = [];
        $source = $dompath->query('//m:cellXfs');
        if (($source->length > 0) && $source->item(0)->hasChildNodes()) {
            foreach ($source->item(0)->childNodes as $id => $style) {
                $styles[$id] = [
                    'node' => $style,
                    'string' => $style->C14N(true, false),
                    'id' => $id,
                ];
            }
        }
        return $styles;
    }

    /**
     * @param \DOMXPath $existing_xpath
     * @param \DOMXPath $source_xpath
     * @param $mapping
     *
     * @return array
     */
    protected function rewriteCells($existing_xpath, $source_xpath, $mapping): array
    {
        $this->definedStyles = $this->getFileStyles($existing_xpath);

        $last_id = \count($this->definedStyles);

        $styles_mapping = [];
        $elems = $source_xpath->query('//m:cellXfs');
        if (($elems->length > 0) && $elems->item(0)->hasChildNodes()) {
            foreach ($elems->item(0)->childNodes as $id => $style) {
                $fontId = (int)$style->getAttribute('fontId');
                if (array_key_exists($fontId, $mapping['fonts'])) {
                    $style->setAttribute('fontId', 0 + $mapping['fonts'][$fontId]);
                }

                $numFmtId = (int)$style->getAttribute('numFmtId');
                if (array_key_exists($numFmtId, $mapping['numFmts'])) {
                    $style->setAttribute('numFmtId', 0 + $mapping['numFmts'][$numFmtId]);
                }

                $fillId = (int)$style->getAttribute('fillId');
                if (array_key_exists($fillId, $mapping['fills'])) {
                    $style->setAttribute('fillId', 0 + $mapping['fills'][$fillId]);
                }

                $borderId = (int)$style->getAttribute('borderId');
                if (array_key_exists($borderId, $mapping['borders'])) {
                    $style->setAttribute('borderId', 0 + $mapping['borders'][$borderId]);
                }

                $xfId = (int)$style->getAttribute('xfId');
                if (array_key_exists($xfId, $mapping['cellStyles'])) {
                    $style->setAttribute('xfId', 0 + $mapping['cellStyles'][$xfId]);
                }

                $string = $style->C14N(true, false);

                foreach ($this->definedStyles as $d) {
                    if ($d['string'] === $string) {
                        // we found an existing style
                        $styles_mapping[$id] = $d['id'];
                        continue 2;
                    }
                }

                // this is a new style!
                $new_id = $last_id++;
                $this->definedStyles[$new_id] = [
                    'node' => $style,
                    'string' => $style->C14N(true, false),
                    'id' => $new_id,
                ];

                $styles_mapping[$id] = $new_id;
            }
        }

        return [$this->definedStyles, $styles_mapping];
    }

    /**
     * @param $defined_styles
     * @param \DOMXPath $existing_xpath
     */
    protected function replaceStylesList($defined_styles, $existing_xpath): void
    {
        $elems = $existing_xpath->query('//m:cellXfs');
        if ($elems->length > 0) {
            $elem = $elems->item(0);
            while ($elem->hasChildNodes()) {
                $elem->removeChild($elem->firstChild);
            }
            foreach ($defined_styles as $s) {
                $elem->appendChild($existing_xpath->document->importNode($s['node'], true));
            }
            $elem->setAttribute('count', \count($defined_styles));
        }
    }
}
