# 🖥️ O365-Suporte — Ferramenta de Suporte Microsoft 365

> Menu interativo em PowerShell para gerenciamento completo do Microsoft 365 — usuários, licenças, grupos, segurança, Exchange Online e SharePoint Online.

---

## 📋 Sumário

- [Sobre o Projeto](#-sobre-o-projeto)
- [Funcionalidades](#-funcionalidades)
- [Requisitos](#-requisitos)
- [Instalação](#-instalação)
- [Configuração](#-configuração)
- [Como Usar](#-como-usar)
- [Estrutura de Menus](#-estrutura-de-menus)
- [Segurança](#-segurança)
- [Logs e Auditoria](#-logs-e-auditoria)
- [Contribuindo](#-contribuindo)
- [Licença](#-licença)

---

## 💡 Sobre o Projeto

O **O365-Suporte** nasceu da necessidade de centralizar e automatizar tarefas rotineiras do suporte técnico em ambientes Microsoft 365. Em vez de navegar por múltiplos portais administrativos para cada operação, o script oferece um menu interativo unificado que executa as ações mais comuns com segurança, velocidade e rastreabilidade.

Desenvolvido 100% em **Windows PowerShell 5.1**, sem dependência de PowerShell 7 ou ferramentas de terceiros.

---

## ✅ Funcionalidades

### 👤 Gestão de Usuários
- Criar novo usuário com licença e grupo atribuídos automaticamente
- Redefinir senha (padrão ou geração aleatória segura com criptografia)
- Bloquear e desbloquear acesso
- Listar usuários com filtros (todos, ativos, bloqueados, sem licença)
- Alterar nome de exibição ou UPN
- Consultar detalhes completos (licenças, grupos, último login, status MFA)
- Excluir usuário com dupla confirmação

### 📋 Gestão de Licenças
- Atribuir licença a usuário existente
- Retirar licenças com seleção individual ou remoção total
- Relatório de licenças do tenant (contratadas, consumidas e disponíveis)
- Listar usuários sem licença atribuída

### 👥 Gestão de Grupos
- Adicionar usuário a grupo
- Remover usuário de grupo
- Listar grupos e membros com exportação CSV

### 🔐 Segurança e Conformidade
- Revogar todas as sessões ativas de um usuário
- Forçar MFA no próximo login
- Relatório de usuários sem MFA configurado
- Identificar usuários inativos por período definido

### 📧 Exchange Online
- Criar caixa de e-mail compartilhada
- Gerenciar permissões (FullAccess e SendAs)
- Criar sala de reunião ou equipamento
- Configurar redirecionamento de e-mail

### 🔄 Onboarding Completo
Fluxo automatizado em uma única operação:
1. Cria o usuário
2. Atribui licença
3. Adiciona ao grupo geral
4. Adiciona a grupos adicionais (seleção interativa)
5. Envia e-mail de boas-vindas com credenciais

### 🚪 Offboarding Completo
Fluxo automatizado com dupla confirmação:
1. Bloqueia a conta
2. Revoga todas as sessões ativas
3. Remove todas as licenças
4. Remove de todos os grupos
5. Configura resposta automática de ausência
6. Redireciona e-mails para o gestor

### ☁️ SharePoint Online
- Visão geral do tenant com barra de uso de armazenamento
- Listar todos os sites (ordenar por uso, nome ou data)
- Ranking dos maiores consumidores de espaço
- Detalhes completos de um site
- Alterar quota de armazenamento
- Bloquear e desbloquear sites (ReadOnly ou NoAccess)
- Gerenciar proprietários e administradores
- Identificar sites inativos por período
- Relatório de uso do OneDrive por usuário
- Excluir site (move para lixeira — recuperável por 93 dias)
- Configurar política de compartilhamento externo
- Relatório completo de armazenamento com exportação CSV

### 📊 Relatórios e Auditoria
- Log de auditoria com timestamp, operador e ação executada
- Relatório de auditoria filtrável por período (hoje, 7, 30 dias ou tudo)
- Relatório de usuários criados recentemente
- Exportação CSV disponível em todas as listagens

---

## 📦 Requisitos

| Requisito | Versão mínima |
|---|---|
| Windows PowerShell | 5.1 |
| Microsoft.Graph | 2.0.0 |
| ExchangeOnlineManagement | 3.0.0 |
| Microsoft.Online.SharePoint.PowerShell | 16.0.0 |

> **Os módulos são instalados automaticamente na primeira execução do script.**

### Permissões necessárias no tenant
- Administrador Global, **ou**
- Administrador de Usuários + Administrador do SharePoint + Administrador do Exchange

---

## 🚀 Instalação

### 1. Clone o repositório

```powershell
git clone https://github.com/SEU_USUARIO/o365-suporte.git
cd o365-suporte
```

### 2. Libere a execução de scripts (uma única vez)

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 3. Execute o script

```powershell
.\O365-Suporte.ps1
```

Na primeira execução, os módulos necessários serão instalados automaticamente. Após a instalação, uma janela de autenticação será aberta no navegador para login com a conta administrativa.

---

## ⚙️ Configuração

Antes de usar, edite o bloco `CONFIG` no início do script com as informações da sua empresa:

```powershell
$Script:CONFIG = @{
    NomeEmpresa        = 'Sua Empresa'              # Nome da empresa (usado em e-mails e banner)
    SenhaPadrao        = 'SuaSenha@2024'            # Senha temporária para novos usuários
    GrupoGeralEmail    = 'geral@suaempresa.com.br'  # E-mail do grupo que todos os usuários recebem
    PaisDefault        = 'BR'                        # Código do país (BR, US, PT...)
}
```

> **Atenção:** Nunca suba o script para o GitHub com a senha real preenchida. Use um placeholder e defina a senha localmente.

---

## 🖱️ Como Usar

Ao executar o script, o menu principal será exibido:

```
  +============================================+
  |            MENU PRINCIPAL                 |
  +============================================+
  |                                           |
  |  [1]  Gestao de Usuarios                 |
  |  [2]  Gestao de Licencas                 |
  |  [3]  Gestao de Grupos                   |
  |  [4]  Seguranca e Conformidade           |
  |  [5]  Caixas de E-mail e Recursos        |
  |  [6]  Relatorios e Exportacao            |
  |                                           |
  |  [7]  Offboarding Completo               |
  |  [8]  Onboarding Completo                |
  |  [9]  Gestao do SharePoint               |
  |                                           |
  |  [0]  Sair                               |
  +============================================+
```

Navegue digitando o número da opção desejada e pressione `ENTER`. Cada submenu possui opção `[0]` para voltar ao menu anterior.

---

## 🗂️ Estrutura de Menus

```
O365-Suporte.ps1
├── [1] Gestao de Usuarios
│   ├── Criar novo usuario
│   ├── Redefinir senha
│   ├── Bloquear acesso
│   ├── Desbloquear acesso
│   ├── Listar usuarios (filtros)
│   ├── Alterar nome ou UPN
│   ├── Detalhes completos
│   └── Excluir usuario
├── [2] Gestao de Licencas
│   ├── Atribuir licenca
│   ├── Retirar licencas
│   ├── Relatorio de licencas
│   └── Usuarios sem licenca
├── [3] Gestao de Grupos
│   ├── Adicionar ao grupo
│   ├── Remover do grupo
│   └── Listar grupos e membros
├── [4] Seguranca e Conformidade
│   ├── Revogar sessoes ativas
│   ├── Forcar MFA
│   ├── Usuarios inativos
│   └── Usuarios sem MFA
├── [5] Caixas de E-mail e Recursos
│   ├── Criar caixa compartilhada
│   ├── Permissao em caixa compartilhada
│   ├── Criar sala ou equipamento
│   └── Configurar redirecionamento
├── [6] Relatorios e Exportacao
│   ├── Auditoria do script
│   ├── Usuarios criados recentemente
│   ├── Relatorio de licencas
│   ├── Usuarios sem licenca
│   └── Usuarios inativos
├── [7] Offboarding Completo
├── [8] Onboarding Completo
└── [9] Gestao do SharePoint
    ├── Visao geral do tenant
    ├── Listar todos os sites
    ├── Top consumidores de espaco
    ├── Detalhes de um site
    ├── Alterar quota do site
    ├── Bloquear / Desbloquear site
    ├── Gerenciar proprietario
    ├── Sites inativos
    ├── OneDrive por usuario
    ├── Excluir site
    ├── Compartilhamento externo
    └── Relatorio completo de armazenamento
```

---

## 🔒 Segurança

- **Autenticação moderna** via Microsoft Graph com OAuth2 — nenhuma senha é armazenada no script
- **Sessão reutilizada** — o login é solicitado apenas uma vez por execução
- **Dupla confirmação** em operações destrutivas (exclusão de usuário, offboarding, exclusão de site)
- **Geração de senha segura** com `System.Security.Cryptography.RandomNumberGenerator`
- **Escopos mínimos** solicitados no login (apenas o necessário para as operações)
- **Sem dados sensíveis** hardcoded no código

---

## 📁 Logs e Auditoria

O script gera automaticamente dois arquivos de log na pasta `logs\` junto ao script:

| Arquivo | Conteúdo |
|---|---|
| `o365suporte.log` | Todas as operações com timestamp e nível (INFO, AVISO, ERRO, OK) |
| `auditoria.log` | Apenas ações executadas com operador e timestamp |

> **Importante:** Adicione a pasta `logs\` ao `.gitignore` para não expor dados do tenant no repositório.

Exemplo de entrada no log de auditoria:
```
2026-01-15 09:42:31 [AUDITORIA] [admin@empresa.com] ACAO: Criacao de usuario - joao.silva@empresa.com
2026-01-15 10:15:03 [AUDITORIA] [admin@empresa.com] ACAO: Offboarding completo - maria.santos@empresa.com
```

---

## 🤝 Contribuindo

Contribuições são bem-vindas! Para contribuir:

1. Faça um fork do repositório
2. Crie uma branch para sua feature: `git checkout -b feature/nova-funcionalidade`
3. Commit suas alterações: `git commit -m 'Adiciona nova funcionalidade'`
4. Push para a branch: `git push origin feature/nova-funcionalidade`
5. Abra um Pull Request

### Ideias para contribuição
- Suporte a múltiplos tenants
- Interface de configuração interativa
- Integração com Microsoft Teams
- Relatórios em formato HTML
- Testes automatizados com Pester

---

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

---

## ⭐ Se este projeto te ajudou, deixe uma estrela no repositório!

Desenvolvido com PowerShell puro para equipes de suporte Microsoft 365.
