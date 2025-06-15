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

class Prettify
{
    /**
     * Effectue un test de nettoyage sur une chaîne donnée.
     *
     * @param string $str Chaîne d'entrée.
     * @return void
     */
    public function test(string $str): void
    {
        echo $this->removeTags($str);
    }

    /**
     * Trouve la position de la Nième occurrence d'une sous-chaîne.
     */
    private function indexOfN(string $str, string $needle, int $number): int|false
    {
        $pos = 0;
        for ($i = 0; $i < $number; $i++) {
            $pos = strpos($str, $needle, $pos);
            if ($pos === false) { return false; }
            $pos++;
        }
        return $pos - 1;
    }

    /**
     * Trouve un placeholder et le nettoie de ses balises.
     */
    private function findAndReplacePlaceholderWithTags(string $str, int $idx): string|false
    {
        $bracketIdx = strpos($str, '{', $idx);
        if ($bracketIdx === false) { return false; }

        // Nettoyer entre $ et {
        $between = substr($str, $idx + 1, $bracketIdx - $idx - 1);
        if (!empty(trim(strip_tags($between)))) { return false; }

        $str = substr_replace($str, '', $idx + 1, $bracketIdx - $idx - 1);

        // Trouver les accolades
        $bracketIdx = strpos($str, '{', $idx);
        $endBracketIdx = strpos($str, '}', $bracketIdx);
        if ($bracketIdx === false || $endBracketIdx === false)  { return false; }

        $inside = substr($str, $bracketIdx + 1, $endBracketIdx - $bracketIdx - 1);
        $clean = strip_tags($inside);

        // Remplacer le contenu du placeholder
        $str = substr_replace($str, $clean, $bracketIdx + 1, $endBracketIdx - $bracketIdx - 1);

        // Vérifier s’il reste des balises de fermeture collées après }
        $endBracketIdx = strpos($str, '}', $bracketIdx);
        $after = substr($str, $endBracketIdx + 1, 20); // Lire les caractères juste après

        // Supprimer </w:t>, </w:r> s’ils sont présents immédiatement après }
        if (preg_match('/^(<\/[^>]+>)+/', $after, $matches)) {
            $str = substr_replace($str, '', $endBracketIdx + 1, strlen($matches[0]));
        }

        return $str;
    }

    /**
     * Supprime les balises HTML/XML entre les placeholders de type ${...}
     */
    public function removeTags(string $str): string
    {
        $placeholders = substr_count($str, '$');

        for ($i = 0; $i < $placeholders; $i++) {
            $idx = $this->indexOfN($str, '$', $i + 1);
            if ($idx === false) { continue; }

            $cleaned = $this->findAndReplacePlaceholderWithTags($str, $idx);
            if ($cleaned !== false) {
                $str = $cleaned;
            }
        }

        return $str;
    }
}
