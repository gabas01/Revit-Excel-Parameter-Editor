# Revit Excel Parameter Editor

![GitHub repo size](https://img.shields.io/github/repo-size/SEU-USUARIO/Revit-Excel-Parameter-Editor?style=for-the-badge)
![GitHub language count](https://img.shields.io/github/languages/count/SEU-USUARIO/Revit-Excel-Parameter-Editor?style=for-the-badge)
![GitHub top language](https://img.shields.io/github/languages/top/SEU-USUARIO/Revit-Excel-Parameter-Editor?style=for-the-badge)
![GitHub last commit](https://img.shields.io/github/last-commit/SEU-USUARIO/Revit-Excel-Parameter-Editor?style=for-the-badge)

Plugin para Autodesk Revit que permite a edição de parâmetros de elementos em massa através da exportação e importação de tabelas (schedules) com o Microsoft Excel.

![Demonstração do Plugin](https://i.imgur.com/your-image-url.gif)  
*(Sugestão: Grave um GIF curto da tela mostrando o plugin em ação e substitua o link acima)*

---

## 🚀 Funcionalidades

* **Interface Intuitiva:** Uma janela simples e direta dentro do Revit.
* **Seleção de Tabelas:** Carrega automaticamente todas as tabelas (schedules) do seu projeto em uma lista.
* **Exportação Segura:** Exporta os dados da tabela para um arquivo `.xlsx`, garantindo que o `ElementID` de cada objeto seja incluído para uma importação sem erros.
* **Importação Inteligente:** Lê o arquivo Excel, identifica os elementos pelo `ElementID` e atualiza apenas os parâmetros modificados.
* **Transações Seguras:** Todas as modificações são feitas dentro de uma única transação do Revit, permitindo que a operação inteira seja desfeita (`Undo`) com um único clique.

## 📦 Instalação (Para Usuários)

1.  Vá para a página de [**Releases**](https://github.com/SEU-USUARIO/Revit-Excel-Parameter-Editor/releases) deste repositório.
2.  Baixe o arquivo `.zip` da versão mais recente.
3.  Descompacte o arquivo. Você encontrará dois itens principais:
    * A pasta `RevitParameterEditor.bundle` (ou os arquivos `RevitParameterEditor.dll` e `RevitParameterEditor.addin`).
4.  Copie a pasta `RevitParameterEditor.bundle` para a pasta de Add-ins do Revit. O caminho é:
    * `C:\Users\SEU-USUARIO\AppData\Roaming\Autodesk\Revit\Addins\[VERSÃO-DO-REVIT]\`
5.  Inicie o Autodesk Revit. O plugin estará disponível na aba **Suplementos (Add-Ins)**.

## 📋 Como Usar

1.  Com seu projeto aberto no Revit, vá para a aba **Suplementos (Add-Ins)**.
2.  Clique no botão **"Editor de Parâmetros"**.
3.  Na janela do plugin, selecione a tabela que deseja editar na lista.
4.  Clique em **"Exportar para Excel"** e salve o arquivo `.xlsx`.
5.  Abra o arquivo no Excel. **Não modifique a coluna `ElementID`!** Edite os valores nas outras colunas de parâmetros conforme necessário.
6.  Salve suas alterações no Excel.
7.  Volte para o Revit, na janela do plugin, e clique em **"Importar do Excel e Atualizar"**.
8.  Selecione o arquivo Excel que você acabou de modificar.
9.  Uma mensagem de sucesso aparecerá informando quantos elementos foram atualizados. Suas alterações agora estão no modelo do Revit!

## ⚙️ Para Desenvolvedores

Este projeto foi criado com C# usando o .NET Framework 4.8 e a API do Revit 2025.

1.  Clone o repositório:
    ```bash
    git clone [https://github.com/SEU-USUARIO/Revit-Excel-Parameter-Editor.git](https://github.com/SEU-USUARIO/Revit-Excel-Parameter-Editor.git)
    ```
2.  Abra o arquivo de solução (`.sln`) no Visual Studio.
3.  As referências `RevitAPI.dll` e `RevitAPIUI.dll` podem precisar ser realocadas. Aponte-as para os arquivos correspondentes na sua pasta de instalação do Revit.
4.  A dependência da biblioteca `EPPlus` pode ser restaurada via NuGet.
5.  Compile o projeto (Build).

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
