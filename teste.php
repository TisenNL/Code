<?php
// Caminho para o arquivo Excel
$excelFile = 'C:/Users/TisenNL/Documents/Works/eSales/code/Chamado_297899.xls';

// Inicializa o objeto Excel
$excel = new COM("Excel.Application");

// Abre o arquivo Excel
$workbook = $excel->Workbooks->Open($excelFile);

// Seleciona a primeira planilha
$worksheet = $workbook->Sheets(1);

// Define a coluna de onde os dados serão lidos
$column = 4; // Por exemplo, a coluna A

// Define a linha inicial e final para leitura
$startRow = 2; // Supondo que a primeira linha seja o cabeçalho
$endRow = $worksheet->UsedRange->Rows->Count;

// Inicializa um array para armazenar os dados
$data = array();

// Lê os dados da coluna especificada
for ($row = $startRow; $row <= $endRow; $row++) {
    $cellValue = $worksheet->Cells($row, $column)->Value();
    $data[] = $cellValue;
}

// Fecha o arquivo Excel
$workbook->Close(false);
$excel->Quit();

// Libera os objetos COM
$worksheet = null;
$workbook = null;
$excel = null;

// Dividir os dados em lotes de 970 argumentos
$batchSize = 970;
$batches = array_chunk($data, $batchSize);

// Para cada lote, construir e executar uma consulta SQL
$output = ''; // Inicializa uma string para armazenar os resultados
foreach ($batches as $batch) {
    // Convertendo os valores do array em uma string separada por vírgulas
    $placeholders = implode(',', array_map(function($value) {
        return "'$value'"; // Envolve cada valor em aspas simples
    }, $batch));

    // Construindo a cláusula WHERE IN
    $whereInClause = "WHERE NOME_ORIG ($placeholders)";

    // Exemplo de uso na consulta SQL
    $sql = "SELECT * FROM ADMEDI_1484.historico $whereInClause";

    // Adiciona a consulta SQL à string de saída
    $output .= $sql . PHP_EOL . PHP_EOL;
}

// Escreve a string de saída em um arquivo de texto
file_put_contents('output.txt', $output);
?>
