# Automaçao Com Python

Esse sistema serve para automatizar tarefas repetitivas no dia a dia. Funciona como um 'robo'.
Ele está configurado para o sistema que uso no meu trabalho atualmente (HPRO).

Usei a lib 'pandas' para tirar dados do excel que uso no trabalho, e para fazer a automação utilizei a lib 'pyautogui'.
O funcionamento é basicamente o seguinte:
  - Eu alimento meu excel com as informações necessárias para comprar determinado item, como por exemplo: código do item no sistema, quantidade, nome, setor e etc.
  - Depois o pyautogui entra em ação e requisita todos os itens de forma automática (como um robo).

 Leva-se cerca de 1 minuto e 40 segundos para requisitar um item manualmente. Se requisitar 50 itens (o que no meu caso é muito comum) demoraria 1 hora e 20 minutos para requisitar tudo.
 Com o meu 'robo' requisita o item em 4 segundos. Ele demoraria 20 minutos para requisitar os mesmos 50 itens. Ganhando 1 hora em uma tarefa que eu repetia todos os dias.
