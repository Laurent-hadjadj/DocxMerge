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

namespace MaMoulinette\Tests;

use PHPUnit\Framework\TestCase;
use MaMoulinette\DocxMerge;

class DocxMergeTest extends TestCase
{
    private $docxMerge;
    private $testFilesDir;
    private $testOutputDir;
    private static $outputFile = '/output.docx';
    private static $templateFile = '/template.docx';

    protected function setUp(): void
    {
        $this->docxMerge = new DocxMerge();
        $this->testFilesDir = __DIR__ . '/files';
        $this->testOutputDir = __DIR__ . '/temp';
    }

    public function testMergeNoFiles()
    {
        $result = $this->docxMerge->merge([], $this->testOutputDir . static::$outputFile);
        $this->assertEquals(-1, $result);
    }

    public function testMergeInvalidOutputPath()
    {
        $result = $this->docxMerge->merge([$this->testFilesDir . '/file1.docx'], $this->testOutputDir . '/output.txt');
        $this->assertEquals(0, $result);
    }

    public function testMergeSuccess()
    {
        $result = $this->docxMerge->merge([$this->testFilesDir . '/file1.docx', $this->testFilesDir . '/file2.docx'],
            $this->testOutputDir . static::$outputFile);
        $this->assertEquals(0, $result);
    }

    public function testSetValuesTemplateNotExist()
    {
        $result = $this->docxMerge->setValues($this->testFilesDir . '/non_existent_template.docx', $this->testOutputDir . static::$outputFile, []);
        $this->assertEquals(-1, $result);
    }

    public function testSetValuesCannotCreateOutputFile()
    {
        // Mock the copy function to simulate failure
        $copyFunction = function ($source, $dest) {
            return false;
        };

        $result = $this->docxMerge->setValues($this->testFilesDir . static::$templateFile, '/invalid/path/output.docx', [],
$copyFunction);
        $this->assertEquals(-2, $result);
    }

    public function testSetValuesSuccess()
    {
        $data = [
            'name' => 'John Doe',
            'date' => '2023-10-01',
            'tables' => [
                [
                    'item' => ['Item 1', 'Item 2'],
                    'price' => ['10', '20']
                ]
            ]
        ];

        $result = $this->docxMerge->setValues(
            $this->testFilesDir . static::$templateFile,
            $this->testOutputDir . static::$outputFile, $data);
        $this->assertEquals(0, $result);
    }

    public function testSetValuesWithEmptyTable()
    {
        $data = [
            'name' => 'Test User',
            'tables' => [] // Cas de table vide
        ];

        $result = $this->docxMerge->setValues(
            $this->testFilesDir . static::$templateFile,
            $this->testOutputDir . static::$outputFile,
            $data
        );

        $this->assertEquals(0, $result);
        $this->assertFileExists($this->testOutputDir . static::$outputFile);
    }

    /*public function testTemplateDocxNotModified()
    {
        $original = file_get_contents($this->testFilesDir . '/template_original.docx');
        $current = file_get_contents($this->testFilesDir . '/template.docx');

        $this->assertEquals($original, $current, 'Le fichier template.docx a été modifié par un autre test !');
    }*/

    protected function tearDown(): void
    {
        array_map('unlink', glob($this->testOutputDir . '/*.docx'));
    }

}
