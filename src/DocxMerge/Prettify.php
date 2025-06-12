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

namespace DocxMerge\DocxMerge;

/**
 * Class Prettify
 * Handles the prettification of strings, including removing tags and replacing placeholders.
 */
class Prettify
{
    /**
     * Test method to demonstrate the functionality of the Prettify class.
     *
     * @param string $str The input string to be processed.
     * @return void
     */
    public function test(string $str): void
    {
            $this->removeTags($str);
    }

    /**
     * Finds the index of the Nth occurrence of a substring in a string.
     *
     * @param string $str The input string.
     * @param string $needle The substring to search for.
     * @param int $number The occurrence number to find.
     * @return int|false The index of the Nth occurrence, or false if not found.
     */
    private function indexOfN(string $str, string $needle, int $number): int|false
    {
        if (substr_count($str, $needle) < $number) {
                return false;
        }

        $startPos = 0;
        $result = 0;

        for ($i = 0; $i < $number; $i++) {
                $idx = strpos($str, $needle, $startPos);
                if ($idx === false) {
                        return false;
                }
                $result = $idx;
                $startPos = $result + 1;
        }

        return $result;
    }

    /**
     * Finds and replaces placeholders with tags in a string.
     *
     * @param string $str The input string.
     * @param int $idx The starting index to search for placeholders.
     * @return string|false The modified string, or false if no placeholders are found.
     */
    private function findAndReplacePlaceholderWithTags(string $str, int $idx): string|false
    {
        $bracketIdx = strpos($str, '{', $idx);
        if ($bracketIdx === false) {
                return false;
        }

        $space = substr($str, $idx + 1, $bracketIdx - $idx - 1);
        if (!empty($space) && strip_tags($space) !== "") {
                return false;
        }

        $str = substr_replace($str, '', $idx + 1, $bracketIdx - $idx - 1);

        // Refresh bracket index after trying update
        $bracketIdx = strpos($str, '{', $idx);
        if ($bracketIdx === false) {
                return false;
        }

        $endBracketIdx = strpos($str, '}', $bracketIdx);
        if ($endBracketIdx === false) {
                return false;
        }

        $space = substr($str, $bracketIdx + 1, $endBracketIdx - $bracketIdx - 1);
        $placeholderName = strip_tags($space);

        $str = substr_replace($str, $placeholderName, $bracketIdx + 1, $endBracketIdx - $bracketIdx - 1);

        return $str;
    }

    /**
     * Removes tags between placeholders in a string.
     *
     * @param string $str The input string.
     * @return string The modified string with tags removed.
     */
    public function removeTags(string $str): string
    {
        // Remove all tags between placeholders (occur when you delete/backspace in editor)
        $placeholderCandidates = substr_count($str, '$');
        $lastIdx = -1;

        for ($i = 0; $i < $placeholderCandidates; $i++) {
                $idx = $this->indexOfN($str, '$', $i + 1);
                if ($idx === false) {
                        continue;
                }

                $status = $this->findAndReplacePlaceholderWithTags($str, $idx);
                if ($status !== false) {
                        $str = $status;
                }
        }

        return $str;
    }
}
