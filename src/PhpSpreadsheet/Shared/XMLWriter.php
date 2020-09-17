<?php

namespace PhpOffice\PhpSpreadsheet\Shared;

class XMLWriter extends \XMLWriter
{
    public static $debugEnabled = false;

    /** Temporary storage method */
    const STORAGE_MEMORY = 1;
    const STORAGE_DISK = 2;

    /**
     * Temporary filename.
     *
     * @var string
     */
    private $tempFileName = '';

    /**
     * Create a new XMLWriter instance.
     *
     * @param int $pTemporaryStorage Temporary storage location
     * @param string $pTemporaryStorageFolder Temporary storage folder
     */
    public function __construct($pTemporaryStorage = self::STORAGE_MEMORY, $pTemporaryStorageFolder = null)
    {
        // Open temporary storage
        if ($pTemporaryStorage == self::STORAGE_MEMORY) {
            $this->openMemory();
        } else {
            // Create temporary filename
            if ($pTemporaryStorageFolder === null) {
                $pTemporaryStorageFolder = File::sysGetTempDir();
            }
            $this->tempFileName = @tempnam($pTemporaryStorageFolder, 'xml');

            // Open storage
            if ($this->openUri($this->tempFileName) === false) {
                // Fallback to memory...
                $this->openMemory();
            }
        }

        // Set default values
        if (self::$debugEnabled) {
            $this->setIndent(true);
        }
    }

    /**
     * Destructor.
     */
    public function __destruct()
    {
        // Unlink temporary files
        if ($this->tempFileName != '') {
            @unlink($this->tempFileName);
        }
    }

    /**
     * Get written data.
     *
     * @return string
     */
    public function getData()
    {
        if ($this->tempFileName == '') {
            return $this->outputMemory(true);
        }
        $this->flush();

        return file_get_contents($this->tempFileName);
    }

    /**
     * Wrapper method for writeRaw.
     *
     * @param string|string[] $text
     *
     * @return bool
     */
    public function writeRawData($text)
    {
        if (is_array($text)) {
            $text = implode("\n", $text);
        }

        return $this->writeRaw(htmlspecialchars($text));
    }

    public function writeDomElement(\DOMElement $elem)
    {
        $this->startElement($elem->nodeName);

        if ($elem->hasAttributes()) {
            /** @var \DomNode $node */
            foreach ($elem->attributes as $node) {
                $this->writeAttribute($node->nodeName, $node->nodeValue);
            }
        }

        if ($elem->hasChildNodes()) {
            foreach ($elem->childNodes as $elem) {
                if ($elem instanceof \DOMElement) {
                    $this->writeDomElement($elem);
                } elseif ($elem instanceof  \DOMText) {
                    $this->writeRawData($elem->nodeValue);
                }
            }
        }

        $this->endElement();
    }
}
