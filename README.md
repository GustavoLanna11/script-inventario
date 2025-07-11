# script-inventario

<h2>Sobre o projeto ✍</h2>
O repositório armazena um script em python para coletar as informações técnicas de uma máquina e inseri-las numa planilha XLSX, planilha essa que será enviada a um dashboard posteriormente.<br>

<h2>Tecnologias e Bibliotecas 📚</h2>
• Python - linguagem na qual o script foi feito. <br>
• Platform - informações sobre o sistema.  <br>
• Socket - obter nome da máquina. <br>
• os - variáveis ambiente como domínio do Active Directory. <br>
• getpass - capturar usuário logado. <br>
• subprocess - executar comandos como wmic e powershell. <br>
• psutil - coletar informações de memória e disco. <br>
• openpyxl - salvar arquivos em XLSX. <br>


<h2>Como rodar o projeto? 💻</h2>
Antes de iniciar, é necessário que sua máquina tenha uma IDE de desenvolvimento. Python e GIT instalados. Verificado isso, clone o projeto em sua máquina, usando o git bash, com o comando 

```git clone -url repositório-```. <br>

Em seguida, abra o terminal e instale a dependência: <br> 

• ```pip install openpyxl```

<br>
Rode o projeto com: <br>

• ```python inventario_maquina.py``` <br><br>

Se até aqui o processo foi feito corretamente, em seu terminal irá aparecer uma mensagem de confirmação, e a planilha com informações da sua máquina será criada no mesmo diretório que se localiza o script.
