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

use PHPUnit\Framework\TestCase;
use MaMoulinette\Prettify;

/**
 * [Description PrettifyTest]
 */
class PrettifyTest extends TestCase
{
    private static $bonjour = 'Bonjour ${nom}, bienvenue !';

    public function testRemoveTagsWithCleanPlaceholder()
    {
        $prettify = new Prettify();

        $input = static::$bonjour;
        $expected = static::$bonjour;
        $this->assertEquals($expected, $prettify->removeTags($input));
    }

    public function testRemoveTagsWithInjectedTags()
    {
        $prettify = new Prettify();

        $input = 'Bonjour $<w:r><w:t>{no<w:r>m}</w:t></w:r>, bienvenue !';
        $expected = static::$bonjour;
        $this->assertEquals($expected, $prettify->removeTags($input));
    }

    public function testIgnoreInvalidPlaceholders()
    {
        $prettify = new Prettify();

        $input = 'Pas de placeholder ici, juste un $ signe.';
        $expected = $input;
        $this->assertEquals($expected, $prettify->removeTags($input));
    }
}
