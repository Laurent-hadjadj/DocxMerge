<?php

/*
  *  Ma-Moulinette
  *  --------------
  *  Copyright (c) 2021-2025.
  *  Laurent HADJADJ <laurent_h@me.com>.
  *  Licensed Creative Common  CC-BY-NC-SA 4.0.
  *  ---
  *  Vous pouvez obtenir une copie de la licence à l'adresse suivante :
  *  http://creativecommons.org/licenses/by-nc-sa/4.0/
*/

namespace MaMoulinette;

use MaMoulinette\Libraries\TbsZip;

/**
 * Class Docx
 * Handles the manipulation of DOCX files, including merging, adding files, and replacing placeholders.
 */
class Docx
{
  // Path to current docx file
  private string $docxPath;

  // Current _RELS data
  private string $docxRels;
  // Current DOCUMENT data
  private string $docxDocument;
  // Current CONTENT_TYPES data
  private string $docxContentTypes;

  private ?TbsZip $docxZip;

  private string $relsZipPath = "word/_rels/document.xml.rels";
  private string $docZipPath = "word/document.xml";
  private string $contentTypesPath = "[Content_Types].xml";
  private string $altChunkType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
  private string $altChunkContentType = "Application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

  // Array "zip path" => "content"
  private array $headerAndFootersArray = [];

  /**
   * Docx constructor.
   *
   * @param string $docxPath Path to the DOCX file.
   */
  public function __construct(string $docxPath, ?TbsZip $zip = null)
  {
    $this->docxPath = $docxPath;

    $this->docxZip = $zip ?? new TbsZip();
    //$this->docxZip = new TbsZip();
    $this->docxZip->Open($this->docxPath);

    $this->docxRels = $this->readContent($this->relsZipPath);
    $this->docxDocument = $this->readContent($this->docZipPath);
    $this->docxContentTypes = $this->readContent($this->contentTypesPath);

  }

  /**
   * [Description for getDocxRels]
   *
   * @return string
   *
   * Created at: 15/06/2025 11:41:29 (Europe/Paris)
   * @author     Laurent HADJADJ <laurent_h@me.com>
   * @copyright  Licensed Ma-Moulinette - Creative Common CC-BY-NC-SA 4.0.
   */
  public function getDocxRels(): string
  {
      return $this->docxRels;
  }

  /**
   * [Description for getDocxDocument]
   *
   * @return string
   *
   * Created at: 15/06/2025 11:41:32 (Europe/Paris)
   * @author     Laurent HADJADJ <laurent_h@me.com>
   * @copyright  Licensed Ma-Moulinette - Creative Common CC-BY-NC-SA 4.0.
   */
  public function getDocxDocument(): string
  {
      return $this->docxDocument;
  }

  /**
   * [Description for getDocxContentTypes]
   *
   * @return string
   *
   * Created at: 15/06/2025 11:41:34 (Europe/Paris)
   * @author     Laurent HADJADJ <laurent_h@me.com>
   * @copyright  Licensed Ma-Moulinette - Creative Common CC-BY-NC-SA 4.0.
   */
  public function getDocxContentTypes(): string
  {
      return $this->docxContentTypes;
  }


  /**
   * Reads the content of a file in the ZIP archive.
   *
   * @param string $zipPath Path to the file in the ZIP archive.
   * @return string The content of the file.
   */
  private function readContent(string $zipPath): string
  {
    return $this->docxZip->FileRead($zipPath);
  }

  /**
   * Writes content to a file in the ZIP archive.
   *
   * @param string $content The content to write.
   * @param string $zipPath Path to the file in the ZIP archive.
   * @return int Returns 0 on success.
   */
  private function writeContent(string $content, string $zipPath): int
  {
    $this->docxZip->FileReplace($zipPath, $content, TBSZIP_STRING);
    return 0;
  }

  /**
   * Adds a file to the DOCX archive.
   *
   * @param string $filePath Path to the file to add.
   * @param string $zipName Name of the file in the ZIP archive.
   * @param string $refID Reference ID for the file.
   * @param bool $addPageBreak Whether to add a page break before the file.
   * @return void
   */
  public function addFile(string $filePath, string $zipName, string $refID, bool $addPageBreak = false): void
  {
    $content = file_get_contents($filePath);
    $this->docxZip->FileAdd($zipName, $content);

    $this->addReference($zipName, $refID);
    $this->addAltChunk($refID, $addPageBreak);
    $this->addContentType($zipName);
  }

  /**
   * Adds a reference to the DOCX archive.
   *
   * @param string $zipName Name of the file in the ZIP archive.
   * @param string $refID Reference ID for the file.
   * @return void
   */
  public function addReference(string $zipName, string $refID): void //private
  {
    $relXmlString = '<Relationship Target="../' . $zipName . '" Type="' . $this->altChunkType . '" Id="' . $refID . '"/>';

    $p = strpos($this->docxRels, '</Relationships>');
    $this->docxRels = substr_replace($this->docxRels, $relXmlString, $p, 0);
  }

  /**
   * Adds an altChunk to the DOCX archive.
   *
   * @param string $refID Reference ID for the file.
   * @param bool $addPageBreak Whether to add a page break before the file.
   * @return void
   */
  public function addAltChunk(string $refID, bool $addPageBreak): void //private
  {
    $pageBreak = $addPageBreak ? '<w:p><w:r><w:br w:type="page" /></w:r></w:p>' : '';
    $xmlItem = $pageBreak . '<w:altChunk r:id="' . $refID . '"/>';

    $p = strpos($this->docxDocument, '</w:body>');
    $this->docxDocument = substr_replace($this->docxDocument, $xmlItem, $p, 0);
  }

  /**
   * Adds a content type to the DOCX archive.
   *
   * @param string $zipName Name of the file in the ZIP archive.
   * @return void
   */
  public function addContentType(string $zipName): void //private
  {
    $xmlItem = '<Override ContentType="' . $this->altChunkContentType . '" PartName="/' . $zipName . '"/>';

    $p = strpos($this->docxContentTypes, '</Types>');
    $this->docxContentTypes = substr_replace($this->docxContentTypes, $xmlItem, $p, 0);
  }

  /**
   * Loads headers and footers from the DOCX archive.
   *
   * @return void
   */
  public function loadHeadersAndFooters(): void
  {
    $relsXML = new \SimpleXMLElement($this->docxRels);
    foreach ($relsXML as $rel) {
    if ($rel["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" ||
            $rel["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header") {
            $path = "word/" . $rel["Target"];
            $this->headerAndFootersArray[$path] = $this->readContent($path);
      }
    }
  }

  /**
   * Finds and replaces placeholders with styles in the DOCX document.
   *
   * @param string $key The placeholder to search for.
   * @param array $value The value to replace the placeholder with, including styles.
   * @return void
   */
  public function findAndReplaceWithStyles(string $key, array $value): void //private
  {
    $lastPos = 0;
    $positions = [];

    while (($lastPos = strpos($this->docxDocument, $key, $lastPos)) !== false) {
            $positions[] = $lastPos;
            $lastPos += strlen($key);
    }

    foreach ($positions as $position) {
      $wrStartPosition1 = strrpos(substr($this->docxDocument, 0, $position), "<w:r ");
      $wrStartPosition2 = strrpos(substr($this->docxDocument, 0, $position), "<w:r>");

      if ($wrStartPosition1 === false && $wrStartPosition2 === false) {
              continue;
      }
      if ($wrStartPosition1 === false) {
              $wrStartPosition = $wrStartPosition2;
      }
      if ($wrStartPosition2 === false) {
              $wrStartPosition = $wrStartPosition1;
      }
      if ($wrStartPosition1 !== false && $wrStartPosition2 !== false) {
              $wrStartPosition = max($wrStartPosition1, $wrStartPosition2);
      }

      $wrStopPosition = strpos(substr($this->docxDocument, $position), "</w:r>") + $position + 6;

      $wrTagStopPosition = strpos(substr($this->docxDocument, $wrStartPosition), ">") + $wrStartPosition + 1;
      $wrTag = substr($this->docxDocument, $wrStartPosition, $wrTagStopPosition - $wrStartPosition);

      $wPrStartPosition = strpos(substr($this->docxDocument, $wrStartPosition), "<w:rPr") + $wrStartPosition;
      $wPrStopPosition = strpos(substr($this->docxDocument, $wPrStartPosition), "</w:rPr>") + $wPrStartPosition;
      $wPrTag = substr($this->docxDocument, $wPrStartPosition, $wPrStopPosition - $wPrStartPosition);

      $insertString = "";
      $idx = 0;
      $len = count($value);
      foreach ($value as $word) {
          $wPrStyles = $wPrTag;
          foreach ($word["decoration"] as $style) {
            if ($style == "bold") {
              $wPrStyles .= "<w:b/>";
            }
            if ($style == "italic") {
                    $wPrStyles .= "<w:i />";
            }
            if ($style == "underline") {
              $wPrStyles .= '<w:u w:val="single"/>';
            }
          }
          $wPrStyles .= "</w:rPr>";

          $insertPart = $wrTag . $wPrStyles . '<w:t xml:space="preserve">' . $word["value"] . '</w:t></w:r>';
          $insertString .= $insertPart;

          if ($idx != $len - 1) {
            $insertString .= '<w:r><w:t xml:space="preserve"> </w:t></w:r>';
          }

          $idx += 1;
      }

      $this->docxDocument = substr($this->docxDocument, 0, $wrStartPosition) .
      $insertString .substr($this->docxDocument, $wrStopPosition);
    }
  }

  /**
   * Finds and replaces placeholders in the DOCX document.
   *
   * @param string $key The placeholder to search for.
   * @param string|array $value The value to replace the placeholder with.
   * @return void
   */
  public function findAndReplace(string $key, $value): void
  {
    if (is_array($value)) {
            $this->findAndReplaceWithStyles($key, $value);
            return;
    }

    $value = htmlspecialchars($value);

    $this->docxDocument = str_replace($key, $value, $this->docxDocument);
    foreach ($this->headerAndFootersArray as $path => $content) {
            $this->headerAndFootersArray[$path] = str_replace($key, $value, $content);
    }
  }

  /**
   * Finds and replaces the first occurrence of a placeholder in the DOCX document.
   *
   * @param string $key The placeholder to search for.
   * @param string $value The value to replace the placeholder with.
   * @return void
   */
  public function findAndReplaceFirst(string $key, string $value): void
  {
    $pos = strpos($this->docxDocument, $key);
    if ($pos === false) {
            return;
    }

    $leftPart = substr($this->docxDocument, 0, $pos);
    $rightPart = substr($this->docxDocument, $pos + strlen($key));
    $this->docxDocument = $leftPart . $value . $rightPart;
  }

  /**
   * Saves the changes to the DOCX file.
   *
   * @return void
   */
  public function flush(): void
  {
    $this->writeContent($this->docxRels, $this->relsZipPath);
    $this->writeContent($this->docxDocument, $this->docZipPath);
    $this->writeContent($this->docxContentTypes, $this->contentTypesPath);
    foreach ($this->headerAndFootersArray as $path => $content) {
            $this->writeContent($content, $path);
    }

    $tempFile = tempnam(dirname($this->docxPath), "dm");
    $this->docxZip->Flush(TBSZIP_FILE, $tempFile);

    if (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN') {
        $status = copy($tempFile, $this->docxPath);
        if ($status) {
                unlink($tempFile);
        }
    } else {
        rename($tempFile, $this->docxPath);
    }
  }

  /**
   * Copies a row with a placeholder in the DOCX document.
   *
   * @param string $placeholder The placeholder to search for.
   * @param int $n The number of times to copy the row.
   * @return void
   */
  public function copyRowWithPlaceholder(string $placeholder, int $n): void
  {
      $needle = "\${" . $placeholder . "}";
      print_r($needle);

      $pos = strpos($this->docxDocument, $needle);
      if ($pos !== false) {
      $trBeginPos = strrpos(substr($this->docxDocument, 0, $pos), "<w:tr");

      if ($trBeginPos === false) {
        return;
      }

      $trEndPos = strpos(substr($this->docxDocument, $pos), "</w:tr>");

      if ($trEndPos === false) {
        return;
      }

      $trEndPos += strlen("</w:tr>");

      $trBody = substr($this->docxDocument, $trBeginPos, $trEndPos - $trBeginPos);
      $result = "";

      for ($i = 0; $i < $n; $i++) {
        $result .= $trBody;
      }

      $this->docxDocument = substr_replace($this->docxDocument, $result, $trEndPos, 0);
      }
  }

  /**
   * Prepares the DOCX document for processing.
   *
   * @return void
   */
  public function prepare(): void
  {
    $prettify = new Prettify();
    $this->docxDocument = $prettify->removeTags($this->docxDocument);
  }
}
