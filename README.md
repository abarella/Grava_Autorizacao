# GRAVA AUTORIZACAO

## 📋 Descrição

Sistema automatizado para processamento e gerenciamento de Notas Fiscais Eletrônicas (NFe) desenvolvido em Visual Basic .NET. O sistema monitora continuamente arquivos XML de retorno da Secretaria da Fazenda, processa autorizações, cancelamentos, cartas de correção e gera arquivos de impressão para diferentes áreas da empresa.

## 🚀 Funcionalidades Principais

### 1. **Processamento de Autorizações de NFe**
- Monitora arquivos `*-pro-rec.xml` na pasta de retorno
- Extrai dados de protocolo, chave de acesso, status e datas
- Grava informações no banco de dados via stored procedures
- Gera arquivos de impressão específicos por área de emissão

### 2. **Geração de XMLs para Envio**
- Lista NFe pendentes de envio
- Gera arquivos XML no formato padrão da SEFAZ
- Salva arquivos na pasta de envio para processamento

### 3. **Processamento de Cancelamentos**
- Processa arquivos `*-can.xml` e `*-ret-env-canc.xml`
- Grava protocolos de cancelamento no banco
- Move arquivos processados para pasta de arquivamento

### 4. **Geração de Arquivos de Serviços**
- Lista notas de serviço pendentes
- Gera arquivos TXT para almoxarifado
- Implementa sistema de sincronização com servidor remoto

### 5. **Consulta de Situação de NFe**
- Gera XMLs de consulta de situação
- Processa retornos de consulta
- Atualiza status das notas no sistema

### 6. **Carta de Correção Eletrônica (CCe)**
- Gera XMLs de carta de correção
- Processa retornos de CCe
- Cria arquivos de impressão para correções

### 7. **Sistema de Notificação por Email**
- Envia emails automáticos em caso de erros
- Anexa arquivos de erro e rejeição
- Configurável por área de emissão

### 8. **Sistema de Logs**
- Registra todas as operações em arquivo de log
- Captura e registra erros detalhadamente
- Facilita troubleshooting e auditoria

## 🏗️ Arquitetura do Sistema

### **Estrutura de Pastas**
```
C:\Unimake\Uninfe4\
├── Envio\          # XMLs para envio à SEFAZ
├── Retorno\        # XMLs de retorno da SEFAZ
├── Gravados\       # Arquivos processados com sucesso
├── SemProtocolo\   # Arquivos com erro/rejeição
├── Importacao\     # Arquivos para área de importação
├── Pedro\          # Arquivos para almoxarifado
├── CTR\            # Arquivos para área CTR
├── Dirf\           # Arquivos para DIRF
├── Servicos\       # Notas de serviço
├── Erro\           # Arquivos com erro
└── Denegado\       # NFe denegadas
```

### **Bancos de Dados**
- **SGCR**: Sistema de gestão
- **VendasInternet**: Dados de vendas online
- **VendasPelicano**: Sistema de vendas principal
- **IPENFAT**: Sistema de faturamento

## 🔧 Tecnologias Utilizadas

- **Linguagem**: Visual Basic .NET 3.5
- **Framework**: .NET Framework 3.5
- **Banco de Dados**: SQL Server
- **Interface**: Windows Forms
- **Processamento**: Timer automático com threading

## 📁 Estrutura do Projeto

```
GRAVA_AUTORIZACAO_TicoTeco/
├── Class/
│   ├── Conexao.vb          # Gerenciamento de conexões com BD
│   ├── Email.vb            # Sistema de envio de emails
│   └── Executar.vb         # Execução de stored procedures
├── Modules/
│   └── Principal.vb        # Variáveis globais do sistema
├── Images/
│   └── SmartViewCustom.ico # Ícone da aplicação
├── Form1.vb                # Formulário principal
├── Form1.Designer.vb       # Designer do formulário
└── CONFIGURACAO.xml        # Configuração de pastas
```

## ⚙️ Configuração

### **1. Arquivo de Configuração (CONFIGURACAO.xml)**
```xml
<Pastas>
    <Envio>C:\Unimake\Uninfe4\Envio\</Envio>
    <Retorno>C:\Unimake\Uninfe4\Retorno\</Retorno>
    <SemProtocolo>C:\Unimake\Uninfe4\SemProtocolo\</SemProtocolo>
    <Gravados>C:\Unimake\Uninfe4\Gravados\</Gravados>
    <!-- ... outras pastas ... -->
</Pastas>
```

### **2. Conexões de Banco de Dados**
As conexões são configuradas na classe `Conexao.vb`:
- Servidor: TICOTECO (CONTINGÊNCIA) ou UIRAPURU (PRODUÇÃO)
- Usuário: crsa
- Senha: cr9537

### **3. Configuração de Email**
- Servidor SMTP: smtp.ipen.br
- Remetente: nfe@ipen.br

## 🚀 Instalação e Execução

### **Pré-requisitos**
- Windows 7 ou superior
- .NET Framework 3.5 SP1
- SQL Server (versão compatível)
- Acesso às pastas de rede configuradas

### **Instalação**
1. Compile o projeto no Visual Studio
2. Execute o arquivo `GRAVA_AUTORIZACAO.exe`
3. Configure as pastas no arquivo `CONFIGURACAO.xml`
4. Verifique as conexões de banco de dados

### **Execução**
O sistema executa automaticamente em loop contínuo:
- Timer principal: execução a cada ciclo
- Processamento sequencial de arquivos
- Sincronização com servidor remoto
- Logs automáticos de todas as operações

## 📊 Áreas de Emissão Suportadas

--------------------------------------------------------------
| Código | Área                  | Destino dos Arquivos      |
|--------|-----------------------|---------------------------|
| IMP    | Importação/Exportação | `\\10.0.10.79\IMPORTACAO` |
| ALM    | Almoxarifado Pedro    | `\\10.0.10.79\PEDRO`      |
| CTR    | Área CTR              | `\\10.0.10.79\CTR`        |
| CR1    | DIRF (Padrão)         | `\\10.0.10.79\Entrada`    |
| CR2    | DIRF Secundária       | `\\10.0.10.79\Entrada`    |
--------------------------------------------------------------
### **Transportadoras Específicas (DIRF)**
- **BND**: `\\10.0.10.79\BND`
- **REM**: `\\10.0.10.79\REM`
- **HVL**: `\\10.0.10.79\HVL`

## 🔍 Monitoramento e Logs

### **Arquivo de Log**
- Localização: `C:\Unimake\Uninfe4\Backup\log.txt`
- Registra: Data/hora, mensagens de erro, stack trace
- Formato: Texto simples com separadores

### **Exemplo de Log**
```
Data/Hora: 2024-01-15 14:30:25
Mensagem: Erro ao processar arquivo XML
Stack Trace: [detalhes do erro]
--------------------------------------------------
```

## 🛠️ Manutenção

### **Verificações Regulares**
1. **Logs de Erro**: Verificar `log.txt` diariamente
2. **Pasta SemProtocolo**: Revisar arquivos rejeitados
3. **Conexões de Rede**: Testar acesso aos servidores remotos
4. **Espaço em Disco**: Monitorar pastas de processamento

### **Troubleshooting Comum**
- **Arquivos não processados**: Verificar permissões de pasta
- **Erros de conexão**: Validar strings de conexão do BD
- **Emails não enviados**: Verificar configuração SMTP
- **Sincronização falha**: Testar conectividade de rede

## 📈 Performance

### **Otimizações Implementadas**
- **Threading**: Processamento assíncrono de arquivos
- **Lock de Sincronização**: Evita processamento simultâneo
- **Timeouts**: Controle de tempo limite para operações
- **Dispose Pattern**: Gerenciamento adequado de recursos

### **Configurações de Timing**
- Timer principal: 8 segundos entre ciclos
- Sleep entre operações: 200ms - 25s
- Timeout de processamento remoto: 60 segundos

## 🔒 Segurança

### **Controles Implementados**
- Conexões seguras com banco de dados
- Validação de arquivos XML
- Tratamento de exceções robusto
- Logs de auditoria completos

## 🗄️ Stored Procedures Utilizadas

O sistema utiliza as seguintes stored procedures para processamento de dados:

### **Banco VendasInternet**

#### **Processamento de Autorizações**
- `dbo.PNFE_GRAVA_AUTORIZACAO` - Grava dados de autorização de NFe
- `dbo.PNFE_GRAVA_CANCELAMENTO` - Grava protocolos de cancelamento
- `dbo.PNFE_GRAVA_ERRO2` - Registra erros de processamento

#### **Geração de XMLs**
- `dbo.PNFe_GERAXML_NF` - Gera XML da NFe para envio
- `dbo.PNFe_CANCELAMENTO` - Gera XML de cancelamento
- `dbo.PNFe_EventoCancelamento` - Gera XML de evento de cancelamento

#### **Carta de Correção Eletrônica**
- `dbo.PNFe_CartaCorrecao_Lista` - Lista cartas de correção pendentes
- `dbo.PNFe_CartaCorrecao_XML` - Gera XML da carta de correção
- `dbo.PNFe_CartaCorrecao_Autorizacao` - Grava autorização da CCe
- `dbo.PNFe_CartaCorrecao_txt` - Gera arquivo TXT da CCe

### **Banco VendasPelicano**

#### **Listagem e Consultas**
- `dbo.P0110_LISTA_NFE_ENVIAR` - Lista NFe pendentes de envio
- `dbo.P0110_LISTA_NFE_CANCELA` - Lista NFe para cancelamento
- `dbo.P0110_LISTA_NFE_SERVICOS` - Lista notas de serviço
- `dbo.P0110_LISTA_NFE_SITUACAO` - Lista NFe para consulta de situação

#### **Geração de Arquivos de Impressão**
- `dbo.P0110_GERA_NOTA_NFE_GERAL` - Gera arquivo para área geral
- `dbo.P0110_GERA_NOTA_NFE_CTR` - Gera arquivo para área CTR
- `dbo.P0110_GERA_NOTA_NFE_DIRF` - Gera arquivo para DIRF
- `dbo.P0110_GERA_NOTA_NFE_SERVICOS` - Gera arquivo de serviços

### **Banco IPENFAT**

#### **Processamento de NFe Denegadas**
- `dbo.sp_NFS_DENEGADAS` - Processa NFe com status denegado

### **Parâmetros Principais das Procedures**

#### **PNFE_GRAVA_AUTORIZACAO**
- `@recibo_numero` (VARCHAR(20)) - Número do recibo
- `@protocolo_autorizacao` (VARCHAR(20)) - Protocolo de autorização
- `@data_retorno` (DATETIME) - Data de retorno
- `@envio_data` (DATETIME) - Data de envio
- `@data_processamento` (DATETIME) - Data de processamento
- `@envio_lote` (VARCHAR(50)) - Lote de envio
- `@recibo_data` (DATETIME) - Data do recibo
- `@status_nf` (INT) - Status da NFe
- `@chave_acesso` (VARCHAR(50)) - Chave de acesso da NFe

#### **PNFE_GRAVA_CANCELAMENTO**
- `@protocolo_cancelamento` (VARCHAR(20)) - Protocolo de cancelamento
- `@chave_acesso` (VARCHAR(50)) - Chave de acesso da NFe

#### **PNFe_GERAXML_NF**
- `@notafis_oid` (INT) - ID da nota fiscal
- `@nf` (INT) - Número da nota fiscal

#### **P0110_GERA_NOTA_NFE_* (Todas as áreas)**
- `@recibo_numero` (VARCHAR(20)) - Número do recibo/protocolo

## 📞 Suporte

Para suporte técnico ou dúvidas sobre o sistema, consulte:
- Logs do sistema em `C:\Unimake\Uninfe4\Backup\log.txt`
- Documentação das stored procedures do banco de dados
- Configurações de rede e permissões de pasta

---

**Desenvolvido para**: Sistema de Gestão de NFe  
**Por**: Equipe de desenvolvimento SEGDS - Basis
**Versão**: 1.0.0  
**Última Atualização**: 2025 por Alberto Barella Junior
**Motivo**: Alteração de layout
