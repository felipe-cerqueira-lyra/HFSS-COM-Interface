# HFSS-COM-Interface
Interface em Python para a comunicação com o ansys hfss utilizando win32com

# Pre-requisitos
- Testado em python 3.11
- win32com
- pandas
- Ansys HFSS 2019.2 (Para outras versões o código terá que ser alterado)


# Uso
A classe HFSS funciona como um wrapper para a comunicação com o HFSS. A classe possui métodos de alto nível que permitem o controle e automatização das funções mais utilizadas para simulação e obtenção de dados.

Para utilizar basta importar a classe HFSS e instanciá-la, passando como argumento o endereço do projeto que deseja ser aberto
```
from HFSSCOMInterface import HFSS

hfss = HFSS(<path_to_project>)
'''insert code here'''
hfss.close()
```
# Métodos implementados
- Alteração do valor de variáveis;
  - recebe como argumentos um dicionário cujos conjuntos chave/valor são os nomes das variáveis e os valores desejados com unidade ("5mm")
  - caso se trate de uma variável de projeto, o nome deve incluir o símbolo "$"
  - ``` hfss.set_variable({'var1':val1, 'var2':val2, '$var3':val3}) ```
- Edição de materiais;
  - recebe como argumento o nome do matérial a ser editado e um dicionário cujos conjuntos chave/valor são os nomes das propriedades a serem alteradas e os valores desejados
  - ```hfss.edit_material("NomeMaterial", {"permittivity":val1, "permeability":val2, "conductivity":val3, "dielectric_loss_tangent":val4})```
- Simulação;
  - recebe como argumento o nome para simulação
  - ```hfss.analyze('SetupName')```
- Obtenção e/ou exportação das matrizes S, Z e Y;
  - recebe como parâmetros a matriz desejada, o nome da solução no formato "<SetupName>:<Sweep>", e o formato dos dados
  - aceita os formatos "Mag/Pha", "Re/Im" e "db/Pha"
  - ```hfss.get_network_data(<'S'/'Z'/'Y'>,  '<SetupName>:Sweep',<'Mag/Pha'/'Re/Im'/'db/Pha'>)```
  - aceita ".tab", ".m", ".sNp" e ".cit" para exportação
  - ```hfss.export_network_data(<path_to_file>, <'S'/'Z'/'Y'>, '<SetupName>:Sweep',<'Mag/Pha'/'Re/Im'/'db/Pha'>)```
- Obtenção e/ou exportação dos dados de campo distânte e próximos;
  - recebe como parâmetros a medida desejada, o nome da solução no formato "<SetupName>:LastAdaptive"
  - para o campo distânte são necessários a geometria (elevação ou azimute) e a frequência.
  - aceita '.txt', '.csv', '.tab' e '.dat' para exportação
  - ```hfss.get_far_field_data(<'DirTotal'/'GainTotal'/...>, '<SetupName>:LastAdaptive', <'Elevation'/'Azimuth'>, <freq>) ```
  - ```hfss.export_far_field_data(<path_to_file>, '<SetupName>:LastAdaptive', <'DirTotal'/'GainTotal'/...>, <'Elevation'/'Azimuth'>, <freq>)``` 
  - para o campo próximo são necessários o eixo x, a geometria (Near field setup), e a frequência
  - ```hfss.get_near_field_data(<'MaxNearETotal'/'NearETotal'/...>, <'Theta'/'Phi'>, '<SetupName>:LastAdaptive', <geometria>, <freq>) ```
  - ```hfss.export_near_field_data(<path_to_file>, '<SetupName>:LastAdaptive', <'MaxNearETotal'/'NearETotal'>, <'Theta'/'Phi'>,  <geometria>, <freq>)```
- Obtenção e/ou exportação dos parâmetros da antena;
  - recebe como parâmetros a medida desejada, o nome da solução no formato "<SetupName>:LastAdaptive" e a geometria (elevação ou azimute)
  - aceita '.txt', '.csv', '.tab' e '.dat' para exportação
  - ```hfss.get_antenna_parameter_data(<'PeakGain'/'PeakGain'/...>, '<SetupName>:LastAdaptive', <'Elevation'/'Azimuth'>)```
  - ```hfss.export_antenna_parameter_data(<path_to_file>, '<SetupName>:LastAdaptive', <'PeakGain'/'PeakGain'/...>,  <'Elevation'/'Azimuth'>)```
- Limpeza dos cache de solução.
  - limpa os dados armazenados em disco das simulações anteriores. Recomendável chamar com frequência para impedir redução da velocidade das simulações
  - ```hfss.clean_solutions()```

# Restrições
- No momento o código só suporta uma instância. A utilização de multiplas instâncias não foi devidamente testada e terá comportamentos inesperados;
- Modelagem de estruturas;
- Criação de variáveis;
- Suporte somente as geometrias em Azimute e Elevação.

# Work in Progress
- Suporte à multiplas instâncias/paralelização

