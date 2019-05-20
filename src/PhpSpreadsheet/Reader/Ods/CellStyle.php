<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Ods;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Style;

class CellStyle extends AbstractStyle
{
    /**
     * see \PhpOffice\PhpSpreadsheet\Writer\Ods\Content::writeXfStyles.
     *
     * @param \DOMElement $node
     *
     * @throws Exception
     */
    public function readFromDom(\DOMElement $element)
    {
        $styleName = $element->getAttributeNS($this->styleNS, 'name');
        if (isset($this->styleMap[$styleName])) {
            throw new Exception('style name ' . $styleName . ' already in use');
        }

        if (null !== ($styleData = $this->_readFromDom($element))) {
            $this->spreadsheet->addCellXf($styleData);
            $this->styleMap[$styleName] = $styleData->getIndex();

            /** @var \DOMElement $child */
            foreach ($element->childNodes as $child) {
                switch ($child->nodeName) {
                    case 'style:text-properties':
                    case 'text-properties':
                        $this->readTextProperties($child, $styleData);

                        break;

                    case 'style:table-cell-properties':
                    case 'table-cell-properties':
                        $this->readCellProperties($child, $styleData);

                        break;
                }
            }
        }
    }

    protected function readCellProperties(\DOMElement $element, Style $styleData)
    {
        $fill = $styleData->getFill();
        $borders = $styleData->getBorders();

        if (!empty($color = $element->getAttributeNS($this->fontNS, 'background-color'))) {
            if (preg_match('/^#([a-fA-F0-9]{6,8})$/', trim($color), $m)) {
                $fill->setEndColor(new Color($m[1]));
            }
        }

        if (!empty($border = $element->getAttributeNS($this->fontNS, 'border-bottom'))) {
            if (preg_match('/^([0-9\.]+)pt ([^ ]+) #([a-fA-F0-9]{6,8})$/', trim($border), $m)) {
                $borders->getBottom()->setSize($m[1])->setBorderStyle($m[2])->setColor(new Color($m[3]));
            }
        }

        if (!empty($border = $element->getAttributeNS($this->fontNS, 'border-left'))) {
            if (preg_match('/^([0-9\.]+)pt ([^ ]+) #([a-fA-F0-9]{6,8})$/', trim($border), $m)) {
                $borders->getLeft()->setSize($m[1])->setBorderStyle($m[2])->setColor(new Color($m[3]));
            }
        }

        if (!empty($border = $element->getAttributeNS($this->fontNS, 'border-top'))) {
            if (preg_match('/^([0-9\.]+)pt ([^ ]+) #([a-fA-F0-9]{6,8})$/', trim($border), $m)) {
                $borders->getTop()->setSize($m[1])->setBorderStyle($m[2])->setColor(new Color($m[3]));
            }
        }

        if (!empty($border = $element->getAttributeNS($this->fontNS, 'border-right'))) {
            if (preg_match('/^([0-9\.]+)pt ([^ ]+) #([a-fA-F0-9]{6,8})$/', trim($border), $m)) {
                $borders->getRight()->setSize($m[1])->setBorderStyle($m[2])->setColor(new Color($m[3]));
            }
        }
    }

    protected function readTextProperties(\DOMElement $element, Style $styleData)
    {
        $font = $styleData->getFont();

        if ($element->getAttributeNS($this->fontNS, 'font-weight') === 'bold' ||
                $element->getAttributeNS($this->styleNS, 'font-weight-complex') === 'bold' ||
                $element->getAttributeNS($this->styleNS, 'font-weight-asian') === 'bold') {
            $font->setBold(true);
        }

        if ($element->getAttributeNS($this->fontNS, 'font-style') === 'italic') {
            $font->setItalic(true);
        }

        if (!empty($color = $element->getAttributeNS($this->fontNS, 'color'))) {
            if (preg_match('/^#([a-fA-F0-9]{6,8})$/', trim($color), $m)) {
                $font->setColor(new Color($m[1]));
            }
        }

        if (!empty($family = $element->getAttributeNS($this->styleNS, 'font-family'))) {
            $font->setName($family);
        }

        if (!empty($size = $element->getAttributeNS($this->fontNS, 'font-size'))) {
            if (preg_match('/^([0-9\.]+)pt$/', trim($size), $m)) {
                $font->setSize($m[1]);
            }
        }

        if (!empty($underlineStyle = $element->getAttributeNS($this->styleNS, 'text-underline-style')) &&
                Font::UNDERLINE_NONE !== $underlineStyle) {
            //$font->setUnderline($underlineStyle);
            if ($underlineStyle === Font::UNDERLINE_DOUBLE) {
                $font->setUnderline(Font::UNDERLINE_DOUBLE);
            } else {
                $font->setUnderline(Font::UNDERLINE_SINGLE);
            }
        }
    }
}
