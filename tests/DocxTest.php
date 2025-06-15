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

use MaMoulinette\Docx;
use MaMoulinette\Libraries\TbsZip;
use PHPUnit\Framework\TestCase;

class DocxTest extends TestCase
{

    protected string $sampleDocxPath;
    protected string $tempDocxPath;
    protected Docx $docx;
    protected $mockTbsZip;

    protected function setUp(): void
    {
        $this->sampleDocxPath = __DIR__ . '/files/template.docx';
        $this->tempDocxPath = sys_get_temp_dir() . '/template_test_' . uniqid() . '.docx';

        // Crée une copie temporaire du template réel
        copy($this->sampleDocxPath, $this->tempDocxPath);

        // Création du mock TbsZip
        /** @var \PHPUnit\Framework\MockObject\MockObject|TbsZip */
        $this->mockTbsZip = $this->createMock(TbsZip::class);
        $this->mockTbsZip->method('FileRead')->willReturnCallback(function($path) {
            $map = [
                "word/_rels/document.xml.rels" => '<Relationships></Relationships>',
                "word/document.xml" => '<w:document><w:body></w:body></w:document>',
                "[Content_Types].xml" => '<Types></Types>',
            ];
            return $map[$path] ?? null;
        });

        $this->docx = new Docx($this->tempDocxPath, $this->mockTbsZip);
    }

    public function testConstructorLoadsContents()
    {
        $this->assertStringContainsString('<Relationships>', $this->docx->getDocxRels());
        $this->assertStringContainsString('<w:document>', $this->docx->getDocxDocument());
        $this->assertStringContainsString('<Types>', $this->docx->getDocxContentTypes());
    }

    public function testAddReferenceAddsRelationship()
    {
        $this->docx->addReference('word/newfile.docx', 'rId10');
        $rels = $this->docx->getDocxRels();
        $this->assertStringContainsString('Id="rId10"', $rels);
        $this->assertStringContainsString('word/newfile.docx', $rels);
    }

    public function testAddAltChunkAddsChunk()
    {
        $this->docx->addAltChunk('rId10', false);
        $document = $this->docx->getDocxDocument();
        $this->assertStringContainsString('<w:altChunk r:id="rId10"/>', $document);
    }

    public function testAddContentTypeAddsOverride()
    {
        $this->docx->addContentType('word/newfile.docx');
        $contentTypes = $this->docx->getDocxContentTypes();
        $this->assertStringContainsString('<Override ContentType=', $contentTypes);
    }

    public function testFindAndReplaceReplacesText()
    {
        $key = '${name}';
        // Injecte directement dans la propriété pour test (ou prévoir setter dans la classe)
        $reflection = new \ReflectionClass($this->docx);
        $docxDocumentProp = $reflection->getProperty('docxDocument');
        $docxDocumentProp->setAccessible(true);
        $docxDocumentProp->setValue($this->docx, "Hello $key!");

        $this->docx->findAndReplace($key, 'World');

        $docxDocument = $this->docx->getDocxDocument();
        $this->assertStringContainsString('Hello World!', $docxDocument);
    }

    public function testCopyRowWithPlaceholderCopiesRow()
    {
        $placeholder = 'testPlaceholder';
        $rowXml = '<w:tr>Some text ${testPlaceholder}</w:tr>';

        $reflection = new \ReflectionClass($this->docx);
        $docxDocumentProp = $reflection->getProperty('docxDocument');
        $docxDocumentProp->setAccessible(true);
        $docxDocumentProp->setValue($this->docx, 'Start ' . $rowXml . ' End');

        $this->docx->copyRowWithPlaceholder($placeholder, 2);

        $docxDocument = $this->docx->getDocxDocument();
        $count = substr_count($docxDocument, '<w:tr>');
        $this->assertEquals(3, $count); // 1 original + 2 copies
    }

    public function testFlushCallsWriteAndSave()
    {
        $this->mockTbsZip->expects($this->once())
            ->method('Flush')
            ->with(
                $this->equalTo(TBSZIP_FILE),
                $this->callback(function ($path) {
                    // On vérifie juste qu'un fichier temporaire est bien généré
                    $this->assertIsString($path);
                    $this->assertStringContainsString('dm', $path);
                    return true;
                })
            );

            // Appelle réellement la méthode pour que le mock soit activé
        $this->docx->flush();
    }

    protected function tearDown(): void
    {
        if (file_exists($this->tempDocxPath)) {
            unlink($this->tempDocxPath);
        }
    }
}
