#!/bin/bash

# Script para instalação de dependências do sistema (Linux)

echo "Atualizando lista de pacotes..."
sudo apt update

echo "Instalando dependências do sistema..."

# LibreOffice (para conversão de DOCX/XLSX para PDF)
if ! command -v soffice &> /dev/null; then
    echo "Instalando LibreOffice..."
    sudo apt install -y libreoffice
else
    echo "LibreOffice já está instalado."
fi

# Poppler Utils (para pdfunite/pdfmerge)
if ! command -v pdfunite &> /dev/null; then
    echo "Instalando Poppler Utils..."
    sudo apt install -y poppler-utils
else
    echo "Poppler Utils já está instalado."
fi

# Graphviz (para geração de diagramas ER)
if ! command -v dot &> /dev/null; then
    echo "Instalando Graphviz..."
    sudo apt install -y graphviz
else
    echo "Graphviz já está instalado."
fi

echo "Dependências do sistema verificadas/instaladas com sucesso!"
