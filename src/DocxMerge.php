<?php

/*
*   Ma-Moulinette
*   --------------
*   Copyright (c) 2021-2025.
*   Laurent HADJADJ <laurent_h@me.com>.
*   Licensed Creative Common CC-BY-NC-SA 4.0.
*   ---
*   Vous pouvez obtenir une copie de la licence à l'adresse suivante :
*   http://creativecommons.org/licenses/by-nc-sa/4.0/
*/

namespace DocxMerge;

use DocxMerge\DocxMerge\Docx;

class DocxMerge
{
    private static string $docxExtension = '.docx';

    /**
     * Merge files in $docxFilesArray order and create new file $outDocxFilePath.
     *
     * @param array $docxFilesArray Array of DOCX file paths to merge.
     * @param string $outDocxFilePath Path to the output DOCX file.
     * @param bool $addPageBreak Whether to add a page break between files.
     * @return int Returns 0 on success, -1 if no files to merge, -2 if cannot create file.
     */
    public function merge(array $docxFilesArray, string $outDocxFilePath, bool $addPageBreak = false): int
    {
        if (empty($docxFilesArray)) {
                // No files to merge
                return -1;
        }

        if (substr($outDocxFilePath, -5) !== self::$docxExtension) {
                $outDocxFilePath .= self::$docxExtension;
        }

        if (!copy($docxFilesArray[0], $outDocxFilePath)) {
                // Cannot create file
                return -2;
        }

        $docx = new Docx($outDocxFilePath);
        for ($i = 1; $i < count($docxFilesArray); $i++) {
                $docx->addFile($docxFilesArray[$i], "part{$i}" . self::$docxExtension, "rId10{$i}", $addPageBreak);
        }

        $docx->flush();

        return 0;
    }

    /**
     * Replace placeholders in a DOCX template with provided data.
     *
     * @param string $templateFilePath Path to the DOCX template file.
     * @param string $outputFilePath Path to the output DOCX file.
     * @param array $data Array of data to replace placeholders with.
     * @return int Returns 0 on success, -1 if template file does not exist, -2 if cannot create output file.
     */
    public function setValues(string $templateFilePath, string $outputFilePath, array $data): int
    {
        if (!file_exists($templateFilePath)) {
                return -1;
        }

        if (!copy($templateFilePath, $outputFilePath)) {
                // Cannot create output file
                return -2;
        }

        $docx = new Docx($outputFilePath);
        $docx->prepare();
        $docx->loadHeadersAndFooters();

        // Add table rows
        if (array_key_exists("tables", $data)) {
            $firstTable = $data["tables"][0];

            foreach ($firstTable as $key => $value) {
                    // Get first placeholder count (other should be same length)
                    $rowCount = count($firstTable[$key]);

                    // Copy row with specified placeholder N times
                    $docx->copyRowWithPlaceholder($key, $rowCount - 1);
                    break;
            }
        }

        foreach ($data as $key => $value) {
            // Skip table placeholders
            if ($key === "tables") {
                    continue;
            }

            $docx->findAndReplace("\${$key}", $value);
        }

        // Fill tables
        if (array_key_exists("tables", $data)) {
                $firstTable = $data["tables"][0];

                foreach ($firstTable as $key => $valueArray) {
                        foreach ($valueArray as $value) {
                                $docx->findAndReplaceFirst("\${$key}", $value);
                        }
                }
        }

        $docx->flush();

        return 0;
    }
}
