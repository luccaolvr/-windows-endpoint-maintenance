# 🖥️ Endpoint Maintenance Tool

Script de manutenção preventiva e inventário para ambientes Windows corporativos.

Desenvolvido para uso em suporte técnico N1/N2, automatiza diagnóstico, limpeza e coleta de dados de hardware — com registro em planilha Excel compartilhada em rede e cópia automática para o ITSM.

---

## ⚙️ Modos de execução

Ao iniciar, o script pergunta qual modo usar:

| Modo | Tempo estimado | O que faz |
|---|---|---|
| **[1] Rápido** | ~3–5 min | Coleta dados, exibe alertas, grava inventário |
| **[2] Completo** | ~20–45 min | Rápido + limpeza de temporários + remoção de bloatware + otimizações |

---

## 🔍 O que o script coleta

- Hostname, IP, fabricante, modelo e serial
- Processador, RAM (tipo e uso atual), disco (tipo, espaço livre, uso %)
- Sistema operacional e última atualização instalada
- Status do antivírus
- Drivers com erro (filtrando dispositivos virtuais)
- Saúde do disco via SMART
- Uptime da máquina
- Reinicialização pendente (Windows Update)
- Top 5 processos por CPU e RAM
- Eventos críticos no log do sistema (últimas 24h)
- Periféricos: monitor, teclado, mouse
- Usuário do AD e e-mail corporativo

---

## 🧹 Modo Completo — o que otimiza

- Remove bloatware padrão do Windows (Solitaire, Skype, Xbox, YourPhone, etc.)
- Desativa apps em segundo plano
- Limpa pastas temporárias (`C:\Windows\Temp`, `%LOCALAPPDATA%\Temp`, Prefetch)
- Limpa logs de eventos (Application e System)
- Verifica e cria usuário local de suporte (se não existir)

---

## 📤 Saídas geradas

**Terminal:** Status de saúde com alertas coloridos + top processos + bloco de dados para registro no ITSM.

**Área de transferência:** Bloco de dados do ITSM copiado automaticamente ao fim da execução.

**Relatório `.txt`:** Salvo na pasta de rede configurada com nome no formato:
```
Maintenance_Nome_Sobrenome_2026-06-29_1430.txt
```

**Planilha Excel:** Nova linha gravada automaticamente no arquivo de inventário compartilhado em rede.

---

## 🚨 Alertas automáticos

O script gera alertas visuais (amarelo/vermelho) para:

- Antivírus inativo ou não detectado
- Disco com uso acima de 80% (atenção) ou 85% (crítico)
- RAM com uso acima de 85%
- Drivers com erro no Gerenciador de Dispositivos
- Falha SMART no disco
- Uptime acima de 30 dias sem reinicialização
- Reinicialização pendente do Windows Update
- Processo consumindo mais de 80% da RAM
- Eventos críticos no log do sistema nas últimas 24h

---

## 🔧 Configuração

Antes de usar, edite o bloco `$CONFIG` no início do script:

```powershell
$CONFIG = @{
    NetworkShare = "\\SEU_SERVIDOR\SEU_COMPARTILHAMENTO"
    SubFolder    = "Preventiva"
    ExcelFile    = "Inventario.xlsx"
    ExcelSheet   = "Inventario"
    NetUser      = ""   # deixe em branco para usar sessão atual do Windows
    NetPass      = ""
}
```

> ⚠️ **Nunca commite credenciais reais.** Use o campo `NetUser`/`NetPass` apenas localmente ou via variável de ambiente.

---

## ▶️ Como executar

1. Clique com o botão direito no arquivo `endpoint-maintenance.bat`
2. Selecione **"Executar como administrador"**
3. Escolha o modo (Rápido ou Completo)
4. Preencha os dados do atendimento quando solicitado

O script solicita elevação automaticamente caso não seja iniciado como administrador.

---

## 📋 Requisitos

- Windows 10 ou superior
- PowerShell 5.1+
- Microsoft Excel instalado (necessário para gravação na planilha)
- Acesso de rede ao compartilhamento configurado

---

## 📁 Estrutura do repositório

```
windows-endpoint-maintenance/
├── endpoint-maintenance.bat   # Script principal
└── README.md
```

---

## 🗺️ Roadmap

- [ ] Versão sem dependência do Excel (exportação direto para CSV)
- [ ] Suporte a múltiplos discos
- [ ] Log de execução local como fallback quando a rede estiver indisponível
- [ ] Parâmetro de linha de comando `--quick` e `--full` para automação via Task Scheduler

---

## 👤 Autor

**Lucca Oliveira** — IT Support → Cloud/SRE  
[linkedin.com/in/luccaolvr](https://linkedin.com/in/luccaolvr) • [github.com/luccaolvr](https://github.com/luccaolvr)
