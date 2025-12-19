# Report Engine

Uma engine de relatórios poderosa e flexível para gerar Excel (com templates avançados) e PDF.

## Instalação

Para utilizar o pacote localmente em seu prjeto ex.: `packages/alexmarquezini/report-engine`, você deve adicioná-lo ao seu `composer.json` principal usando um repositório do tipo `path`.

1.  Edite o `composer.json` da raiz do seu projeto:

```json
"repositories": [
    {
        "type": "path",
        "url": "./packages/alexmarquezini/report-engine"
    }
],
"require": {
    "alexmarquezini/report-engine": "@dev"
}
```

2.  Execute `composer update`.

## Uso

### 1. Definir o Relatório

```php
use AlexMarquezini\ReportEngine\ReportDefinition;

$definition = new ReportDefinition();
$definition->setTitle("Meu Relatório");
$definition->setColumns([
    'NOME' => ['label' => 'Nome', 'width' => 30],
    'VALOR' => ['label' => 'Valor', 'format' => 'currency']
]);
```

### 2. Processar Dados

```php
use AlexMarquezini\ReportEngine\ReportProcessor;

$data = [...]; // Array de objetos ou arrays
$processor = new ReportProcessor($data, $definition);
$processedData = $processor->process();
```

### 3. Gerar Excel

```php
use AlexMarquezini\ReportEngine\ExcelGenerator;

$generator = new ExcelGenerator($processedData, $definition);
$generator->setTemplate('/path/to/template.xlsx'); // Opcional
$spreadsheet = $generator->generate();
```

### 4. Gerar PDF

```php
use AlexMarquezini\ReportEngine\PdfGenerator;

$pdfGen = new PdfGenerator($processedData, $definition);
$pdfContent = $pdfGen->generate();
```

## Requisitos

-   PHP 7.4+
-   PhpSpreadsheet
-   KnpSnappy (wkhtmltopdf)
-   CodeIgniter 4 (Atualmente dependente de helpers globais como `view()`, `ROOTPATH`)
