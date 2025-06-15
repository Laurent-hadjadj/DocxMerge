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

use MaMoulinette\Docx;

/**
 * [Description DocxMerge]
 */
class DocxMerge
{
    private static string $docxExtension = '.docx';

    /**
     * [Description for merge]
     *
     * @param array $docxFilesArray
     * @param string $outDocxFilePath
     * @param bool $addPageBreak
     *
     * @return int
     *
     * Created at: 15/06/2025 11:42:50 (Europe/Paris)
     * @author     Laurent HADJADJ <laurent_h@me.com>
     * @copyright  Licensed Ma-Moulinette - Creative Common CC-BY-NC-SA 4.0.
     */
    public function merge(array $docxFilesArray, string $outDocxFilePath, bool $addPageBreak = false): int
    {
        if (empty($docxFilesArray)) {
            // No files to merge
            return -1;
        }

        if (!str_ends_with($outDocxFilePath, self::$docxExtension)) {
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
     * [Description for setValues]
     *
     * @param string $templatePath
     * @param string $outputPath
     * @param array $data
     * @param callable|null $copyFunction
     *
     * @return int
     *
     * Created at: 15/06/2025 11:42:56 (Europe/Paris)
     * @author     Laurent HADJADJ <laurent_h@me.com>
     * @copyright  Licensed Ma-Moulinette - Creative Common CC-BY-NC-SA 4.0.
     */
    public function setValues(string $templatePath, string $outputPath, array $data, ?callable $copyFunction = null): int
    {
        // Vérifier l'existence du template
        if (!file_exists($templatePath)) {
            return -1; // Template non trouvé
        }

        // Vérifier que l'extension de sortie est .docx
        if (pathinfo($outputPath, PATHINFO_EXTENSION) !== 'docx') {
            return 0; // Mauvais format de sortie
        }

        // Copier le fichier template vers la destination
        if ($copyFunction === null) {
            $copyFunction = function ($src, $dst) {
                return copy($src, $dst);
            };
        }

        if (!$copyFunction($templatePath, $outputPath)) {
            return -2; // Impossible de copier le fichier
        }

        // Charger le document
        $docx = new Docx($outputPath);

        // Remplir les champs simples
        foreach ($data as $key => $value) {
            if ($key === 'tables') {
                // On gère les tables plus bas
                continue;
            }

            if (is_string($value) || is_numeric($value)) {
                $docx->findAndReplace("\${$key}", (string)$value);
            }
        }

        // Gestion des tables (si elles existent et ne sont pas vides)
        if (
            isset($data["tables"]) &&
            is_array($data["tables"]) &&
            count($data["tables"]) > 0 &&
            isset($data["tables"][0]) &&
            is_array($data["tables"][0]) &&
            count($data["tables"][0]) > 0
        ) {
            $firstTable = $data["tables"][0];

            // Copier les lignes avec placeholders
            foreach ($firstTable as $key => $valueArray) {
                if (is_array($valueArray)) {
                    $rowCount = count($valueArray);
                    $docx->copyRowWithPlaceholder($key, $rowCount - 1);
                    break; // On ne copie qu'une fois
                }
            }

            // Remplir les tableaux
            foreach ($data["tables"] as $table) {
                if (!is_array($table))  { continue; }

                foreach ($table as $key => $valueArray) {
                    if (!is_array($valueArray)) { continue; }

                    foreach ($valueArray as $value) {
                        $docx->findAndReplaceFirst("\${$key}", $value);
                    }
                }
            }
        }

        // Sauvegarder le document modifié
        $docx->flush();

        return 0; // Succès
    }
}
