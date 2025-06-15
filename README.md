# DocxMerge

┏┓┏┓┏┓━┏━━┓━┏━━━┓\
┃┃┃┃┃┃━┗┫┣┛━┃┏━┓┃\
┃┃┃┃┃┃━━┃┃━━┃┗━┛┃\
┃┗┛┗┛┃━━┃┃━━┃┏━━┛\
┗┓┏┓┏┛━┏┫┣┓━┃┃\
━┗┛┗┛━━┗━━┛━┗┛  NeXt 1.1.0

🟢 **Version** : 1.1.0-Release\
🔁 **Fork** de DocxMerge 1.0.1 (<https://github.com/krustnic/DocxMerge/>)\
📥 **GitHub** : <https://github.com/Laurent-hadjadj/DocxMerge/>

## Introduction

Simple bibliothèque pour fusionner plusieurs fichiers MS Word `.docx` en un seul.

Cette librairie date de 2016 pour la dernière version 1.0.1. Plusieurs propositions de corrections sont disponibles sur le dépôt GitHub d’origine, mais aucune n’a été fusionnée à ce jour.

Nous utilisons cette librairie dans notre projet, conjointement avec PhpWord, pour fusionner un modèle de document avec un fichier généré dynamiquement. 

Nous avons donc décidé de forker la version 1.0.1 et d’y intégrer les corrections nécessaires. Nous en avons également profité pour mettre à jour la version de TbsZip afin d’assurer la compatibilité avec PHP 8.2 et supérieur.

> **Note :**
> Les modifications apportées permettent une utilisation immédiate, mais une réécriture complète du code serait souhaitable pour en améliorer la maintenabilité.

## Licence

### Code source

Ce projet est sous licence Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International. Voir le fichier [LICENSE-CC](LICENSE-CC) pour plus de détails.

### Package

Ce package est sous licence MIT. Voir le fichier [LICENSE](LICENSE) pour plus de détails.

## Fonctionnalités

- Génère un fichier DOCX valide pour MS Office 2007 et versions ultérieures

## Détails techniques

Pour manipuler l’archive ZIP du DOCX, nous utilisons [TbsZip](http://www.tinybutstrong.com/MaMoulinettes/tbszip/tbszip_help.html).

## Installation

```bash
composer require laurent-hadjadj/docx-merge
```

## Exemple de fusion (merge)

```php
require 'vendor/autoload.php';
use MaMoulinette\DocxMerge;

$dm = new DocxMerge();
$dm->merge(
    [
        'templates/TplPage1.docx',
        'templates/TplPage2.docx',
    ],
    '/tmp/result.docx'
);
```

## Exemple de remplacement de valeurs (setValues)

> Utilisez des placeholders comme ${NAME} dans votre fichier `.docx`.

```php
require 'vendor/autoload.php';
use MaMoulinette\DocxMerge;

$dm = new DocxMerge();
$dm->setValues(
    'templates/template.docx',
    'templates/result.docx',
    [
        'NAME'    => 'Sterling',
        'SURNAME' => 'Archer',
    ]
);
```

### Avec styles (gras, italique, souligné)

```php
$dm->setValues(
    'templates/template.docx',
    'templates/result.docx',
    [
        'NAME' => [
            [
                'value'      => 'Sterling',
                'decoration' => ['bold', 'italic'],
            ],
            [
                'value'      => 'Archer',
                'decoration' => ['bold', 'underline'],
            ],
        ],
    ]
);
```
