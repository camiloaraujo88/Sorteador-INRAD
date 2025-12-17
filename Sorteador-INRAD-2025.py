import pygame
import random
import sys
import pandas as pd
from datetime import datetime
import os
from pygame import mixer
import openpyxl

# Inicializa o Pygame e o mixer de áudio
pygame.init()
mixer.init()

# Configurações da tela
LARGURA, ALTURA = 1200, 800
tela = pygame.display.set_mode((LARGURA, ALTURA), pygame.RESIZABLE)
pygame.display.set_caption("Sorteador de Confraternização")

# Variáveis globais
tela_cheia = False
audio_ativado = True
largura_tela, altura_tela = LARGURA, ALTURA

# Lista global de categorias (com repetições conforme quantidades desejadas)
# Comentários (PT-BR):
# - Mantém acentuação exatamente como na planilha para casar corretamente com a chave 'categoria'.
# - Riscos: Se houver divergências de grafia/acentuação nas categorias da planilha, pode não haver correspondência exata.
# - Sugestão: Padronize as categorias na fonte de dados. Caso alguma categoria não exista entre os participantes,
#   a cota ainda será consumida, o que pode alterar a distribuição real. Ajuste conforme a necessidade do evento.
categorias_participantes = (
    ['ADM+APOIO'] * 6 +
    ['HRB'] * 1 +
    ['MÉDICO'] * 4 +
    ['MULTI'] * 11 +
    ['RESIDENTES'] * 5 +
    ['TERCEIROS'] * 4
)

# Obtém o caminho da área de trabalho do usuário
CAMINHO_DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")

# Caminhos dos arquivos na área de trabalho
CAMINHO_EXCEL = os.path.join(CAMINHO_DESKTOP, "sorteio.xlsx")
CAMINHO_LOG = os.path.join(CAMINHO_DESKTOP, "log_sorteios_inrad.txt")

print(f"📁 Procurando planilha em: {CAMINHO_EXCEL}")
print(f"📁 Log será salvo em: {CAMINHO_LOG}")

# Verifica se estamos em um executável PyInstaller
def is_exe():
    """Verifica se o script está rodando como executável"""
    return hasattr(sys, '_MEIPASS')

# Função para obter o caminho correto dos recursos
def resource_path(relative_path):
    """Obtém o caminho absoluto para recursos, funciona para dev e para exe"""
    if is_exe():
        # Se for executável, usa a pasta temporária do PyInstaller
        base_path = sys._MEIPASS
    else:
        # Se for desenvolvimento, usa o diretório atual
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# Caminhos dos áudios embutidos no executável
AUDIO_FILES = {
    'tambor': 'tambor.mp3',
    'banheira': 'banheira.mp3', 
    'piao': 'piao.mp3',
    'carlton': 'carlton.mp3',
    'moveyourfeet': 'moveyourfeet.mp3',
    'labamba': 'labamba.mp3',
    'mexe': 'mexe.mp3',
    'balaio': 'balaio.mp3',
    'musicaepica': 'musicaepica.mp3',
    'snoop': 'snoop.mp3',
    'suspense2': 'suspense2.mp3',
    'genius': 'genius.mp3',
    'gta': 'gta.mp3',
    'need': 'need.mp3',
    'macarena': 'macarena.mp3',
    'sucesso': 'sucesso.mp3'
}

# Variáveis globais para áudios
audio_atual_index = 0  # Índice do áudio atual na sequência
audio_atual = None  # Áudio que está tocando atualmente
tempo_inicio_audio = 0
audio_em_execucao = False
duracao_audio = 8000  # 8 segundos em milissegundos

# CORREÇÃO ALTERNATIVA 2: Carrega áudios diretamente do _MEIPASS
def carregar_audios():
    """Carrega áudios embutidos no executável"""
    global audios_sorteio, som_sucesso
    
    audios = {}
    
    for nome, arquivo in AUDIO_FILES.items():
        try:
            if is_exe():
                # No executável, os arquivos estão em sys._MEIPASS
                caminho = os.path.join(sys._MEIPASS, arquivo)
                print(f"🔍 Procurando áudio em (executável): {caminho}")
            else:
                # Para desenvolvimento, procura na pasta 'audio' local
                caminho = os.path.join("audio", arquivo)
                print(f"🔍 Procurando áudio em (dev): {caminho}")
            
            if os.path.exists(caminho):
                som = mixer.Sound(caminho)
                audios[nome] = som
                print(f"✓ Áudio carregado: {arquivo}")
            else:
                print(f"⚠ Áudio não encontrado: {arquivo}")
                print(f"  Caminho testado: {caminho}")
                audios[nome] = None
                
        except Exception as e:
            print(f"✗ Erro ao carregar áudio {arquivo}: {e}")
            # Informações adicionais para debug
            if is_exe():
                print(f"  sys._MEIPASS: {sys._MEIPASS}")
                print(f"  Arquivos em _MEIPASS: {os.listdir(sys._MEIPASS) if os.path.exists(sys._MEIPASS) else 'Não acessível'}")
            audios[nome] = None
    
    # CORREÇÃO: Usando o método simplificado para evitar erros de sintaxe
    audios_sorteio = []
    for nome in ['tambor', 'banheira', 'piao', 'carlton', 'moveyourfeet', 
                 'labamba', 'mexe', 'balaio', 'musicaepica', 'snoop', 
                 'suspense2', 'genius', 'gta', 'need', 'macarena']:
        audios_sorteio.append(audios.get(nome))
    
    som_sucesso = audios.get('sucesso')
    
    # Contagem de áudios carregados
    carregados = sum(1 for a in audios_sorteio if a is not None)
    sucesso_carregado = 1 if som_sucesso else 0
    
    print(f"\n🎵 RESUMO DE ÁUDIOS:")
    print(f"  Áudios de sorteio carregados: {carregados}/15")
    print(f"  Áudio de sucesso carregado: {'Sim' if som_sucesso else 'Não'}")
    print(f"  Sequência com {len(audios_sorteio)} áudios")
    
    return audios_sorteio, som_sucesso

# Carrega os áudios
audios_sorteio, som_sucesso = carregar_audios()

# Cores
COR_FUNDO = (15, 25, 40)
COR_TEXTO = (255, 255, 255)
COR_DESTAQUE = (255, 215, 0)
COR_CATEGORIA = (100, 200, 255)
COR_BOTAO = (70, 130, 180)
COR_BOTAO_HOVER = (100, 160, 210)
COR_BOTAO_SORTEIO = (220, 60, 60)
COR_BOTAO_SORTEIO_HOVER = (255, 90, 90)
COR_VERDE = (60, 180, 75)
COR_AMARELO = (255, 200, 50)
COR_AZUL = (80, 160, 255)
COR_RODAPE = (25, 35, 55)

# Fontes
fonte_grande = pygame.font.SysFont('arial', 48, bold=True)
fonte_media = pygame.font.SysFont('arial', 32, bold=True)
fonte_normal = pygame.font.SysFont('arial', 24)
fonte_pequena = pygame.font.SysFont('arial', 18)
fonte_rodape = pygame.font.SysFont('arial', 22, bold=True)

# Carrega dados do Excel
def carregar_dados_excel():
    try:
        # Verifica se o arquivo existe na área de trabalho
        if not os.path.exists(CAMINHO_EXCEL):
            print("⚠ Planilha não encontrada na área de trabalho!")
            print(f"⚠ Procurei em: {CAMINHO_EXCEL}")
            
            # Tenta encontrar o arquivo com nomes alternativos
            arquivos_alternativos = [
                "CONFRATERNIZCAO-SORTEIO.xlsx",
                "sorteio.xlsx",
                "CONFRATERNIZACAO-SORTEIO.xlsx",
                "SORTEIO.xlsx"
            ]
            
            for arquivo in arquivos_alternativos:
                caminho_alternativo = os.path.join(CAMINHO_DESKTOP, arquivo)
                if os.path.exists(caminho_alternativo):
                    print(f"✓ Encontrei arquivo alternativo: {arquivo}")
                    df = pd.read_excel(caminho_alternativo)
                    print(f"✓ Planilha carregada: {len(df)} participantes")
                    return processar_dataframe(df)
            
            print("⚠ Usando dados de exemplo por enquanto...")
            return criar_dados_exemplo()
        
        df = pd.read_excel(CAMINHO_EXCEL)
        print(f"✓ Planilha carregada: {len(df)} participantes")
        return processar_dataframe(df)
    
    except Exception as e:
        print(f"✗ Erro ao carregar Excel: {e}")
        return criar_dados_exemplo()

def processar_dataframe(df):
    """Processa o dataframe para extrair participantes"""
    # Tenta detectar automaticamente as colunas
    colunas_df = [str(col).upper().strip() for col in df.columns]
    
    print(f"Colunas encontradas: {colunas_df}")
    
    # Procura colunas por padrões comuns
    categoria_idx = None
    id_idx = None
    nome_idx = None
    
    # Procura por diferentes nomes possíveis
    for i, col in enumerate(colunas_df):
        col_clean = col.replace(" ", "").replace("_", "").replace("-", "")
        
        if any(keyword in col_clean for keyword in ['CATEGORIA', 'SETOR', 'DEPARTAMENTO', 'GRUPO', 'TIPO']):
            categoria_idx = i
        elif any(keyword in col_clean for keyword in ['ID', 'MATRICULA', 'CODIGO', 'NÚMERO', 'NUMERO', 'COD']):
            id_idx = i
        elif any(keyword in col_clean for keyword in ['NOME', 'NOMES', 'PARTICIPANTE', 'FUNCIONARIO', 'PESSOA']):
            nome_idx = i
    
    # Se não encontrou, usa lógica de fallback
    if categoria_idx is None:
        # Tenta encontrar pelo conteúdo da primeira linha
        try:
            primeira_linha = df.iloc[0]
            for i, valor in enumerate(primeira_linha):
                if isinstance(valor, str) and any(keyword in valor.upper() for keyword in ['FUNC', 'FORN', 'CLI', 'DIR']):
                    categoria_idx = i
                    break
        except:
            pass
        
        if categoria_idx is None:
            categoria_idx = 0
    
    if id_idx is None:
        id_idx = 1 if len(df.columns) > 1 else categoria_idx
    
    if nome_idx is None:
        # Procura por coluna que pareça ter nomes
        for i in range(len(df.columns)):
            if i != categoria_idx and i != id_idx:
                nome_idx = i
                break
        
        if nome_idx is None:
            nome_idx = 2 if len(df.columns) > 2 else 0
    
    print(f"Usando colunas: Categoria[{categoria_idx}], ID[{id_idx}], Nome[{nome_idx}]")
    
    participantes = []
    for _, row in df.iterrows():
        try:
            cat = str(row.iloc[categoria_idx]) if categoria_idx < len(df.columns) else "Geral"
            id_val = str(row.iloc[id_idx]) if id_idx < len(df.columns) else "N/A"
            nome = str(row.iloc[nome_idx]) if nome_idx < len(df.columns) else "Desconhecido"
            
            # Limpa os valores
            cat = cat.strip()
            id_val = str(id_val).strip()
            nome = nome.strip()
            # NOVO: Força o nome para MAIÚSCULAS ao carregar do dataframe (pedido do usuário)
            # Comentários (PT-BR):
            # - Isso padroniza a exibição dos nomes na animação e no resultado final.
            # - Risco/Sugestão: pode aumentar a largura do texto (nomes em CAPS ocupam mais espaço);
            #   o layout já usa texto responsivo para se ajustar, mas nomes muito longos podem reduzir o tamanho da fonte.
            nome = nome.upper()
            
            # Pula linhas vazias
            if not nome or nome == "nan" or nome == "None":
                continue
                
            participantes.append({
                'categoria': cat,
                'id': id_val,
                'nome': nome
            })
        except Exception as e:
            print(f"✗ Erro ao processar linha: {e}")
            continue
    
    print(f"✓ Total de participantes processados: {len(participantes)}")
    return participantes

def criar_dados_exemplo():
    """Cria dados de exemplo quando não há planilha"""
    print("⚠ Criando dados de exemplo...")
    
    # Cria um arquivo de exemplo na área de trabalho
    exemplo_excel = os.path.join(CAMINHO_DESKTOP, "EXEMPLO-SORTEIO.xlsx")
    try:
        dados_exemplo = {
            'Categoria': ['Funcionários', 'Funcionários', 'Fornecedores', 'Fornecedores', 
                         'Clientes', 'Clientes', 'Diretoria', 'Diretoria'],
            'ID': ['001', '002', '003', '004', '005', '006', '007', '008'],
            'Nome': ['João Silva', 'Maria Santos', 'Carlos Mendonça', 'Ana Pereira',
                    'Pedro Costa', 'Juliana Rodrigues', 'Roberto Almeida', 'Fernanda Lima']
        }
        df_exemplo = pd.DataFrame(dados_exemplo)
        df_exemplo.to_excel(exemplo_excel, index=False)
        print(f"✓ Arquivo de exemplo criado: {exemplo_excel}")
    except Exception as e:
        print(f"✗ Não foi possível criar arquivo de exemplo: {e}")
    
    return [
        {'categoria': 'Funcionários', 'id': '001', 'nome': 'João Silva'},
        {'categoria': 'Funcionários', 'id': '002', 'nome': 'Maria Santos'},
        {'categoria': 'Fornecedores', 'id': '003', 'nome': 'Carlos Mendonça'},
        {'categoria': 'Fornecedores', 'id': '004', 'nome': 'Ana Pereira'},
        {'categoria': 'Clientes', 'id': '005', 'nome': 'Pedro Costa'},
        {'categoria': 'Clientes', 'id': '006', 'nome': 'Juliana Rodrigues'},
        {'categoria': 'Diretoria', 'id': '007', 'nome': 'Roberto Almeida'},
        {'categoria': 'Diretoria', 'id': '008', 'nome': 'Fernanda Lima'},
    ]

# Função para salvar log na área de trabalho
def salvar_log_sorteio(participante):
    try:
        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        log_entry = f"{data_hora} | Categoria: {participante['categoria']} | ID: {participante['id']} | Nome: {participante['nome']}\n"
        
        # Verifica se a área de trabalho existe
        if not os.path.exists(CAMINHO_DESKTOP):
            os.makedirs(CAMINHO_DESKTOP, exist_ok=True)
            print(f"📁 Criada pasta da área de trabalho: {CAMINHO_DESKTOP}")
        
        with open(CAMINHO_LOG, 'a', encoding='utf-8') as f:
            f.write(log_entry)
        
        print(f"✓ Log salvo na área de trabalho: {participante['nome']}")
        return True
    except Exception as e:
        print(f"✗ Erro ao salvar log: {e}")
        return False

class Botao:
    def __init__(self, x, y, largura, altura, texto, cor_normal=COR_BOTAO):
        self.rect = pygame.Rect(x, y, largura, altura)
        self.texto = texto
        self.cor_normal = cor_normal
        self.cor = cor_normal
        self.cor_hover = self.calcular_cor_hover(cor_normal)
        
    def calcular_cor_hover(self, cor):
        return tuple(min(c + 40, 255) for c in cor)
        
    def desenhar(self, superficie):
        pygame.draw.rect(superficie, self.cor, self.rect, border_radius=10)
        pygame.draw.rect(superficie, (255, 255, 255), self.rect, 2, border_radius=10)
        
        fonte_botao = pygame.font.SysFont('arial', 22)
        texto_surf = fonte_botao.render(self.texto, True, COR_TEXTO)
        texto_rect = texto_surf.get_rect(center=self.rect.center)
        superficie.blit(texto_surf, texto_rect)
        
    def verificar_clique(self, pos):
        return self.rect.collidepoint(pos)
        
    def verificar_hover(self, pos):
        self.cor = self.cor_hover if self.rect.collidepoint(pos) else self.cor_normal
        return self.rect.collidepoint(pos)

class Sorteador:
    def __init__(self):
        self.participantes = carregar_dados_excel()
        self.sorteando = False
        self.participante_sorteado = None
        self.historico = []
        self.tempo_sorteio = 0
        self.velocidade_sorteio = 30
        self.tempo_finalizacao = 0
        self.animacao_atual = 0
        self.participantes_sorteados_ids = set()
        self.ultimo_sorteio = ""
        
        # Contador de áudios usado para sequência
        self.contador_sorteios = 0  # Conta quantos sorteios já foram feitos
        self.audio_sorteio_atual = None  # Áudio do sorteio atual
        
        # Controle do pool de categorias para cumprir as cotas por rodada
        # Comentários (PT-BR): criamos um backup fixo e um pool mutável para consumo a cada sorteio.
        # Risco: se a planilha não tiver participantes suficientes em uma categoria para a cota, haverá fallback (ver método sortear_participante)
        self.categorias_pool_backup = list(categorias_participantes)  # cópia imutável da configuração inicial
        self.categorias_pool = None  # será inicializada no primeiro sorteio
        
        self.atualizar_botoes()
        
    def atualizar_botoes(self):
        """Atualiza posições dos botões"""
        global largura_tela, altura_tela
        
        # Botões principais
        btn_largura = 250
        btn_altura = 55
        espacamento = 30
        
        # Posição Y: mais para cima (sem a barra de progresso)
        pos_y = altura_tela - 180  # Subiu 40px
        
        # Centraliza o único botão disponível (removido o botão "NOVO SORTEIO")
        inicio_x = (largura_tela - btn_largura) // 2  # centralizado
        
        self.botao_sortear = Botao(inicio_x, pos_y, btn_largura, btn_altura, 
                                   "SORTEAR", COR_BOTAO_SORTEIO)
    
    def sortear_participante(self):
        """Sorteia um participante que ainda não foi sorteado"""
        # Checagem de nulos/inconsistências na lista de participantes
        if not self.participantes:
            # Aviso: sem participantes carregados – verifique a planilha
            return None
        
        # Inicializa o pool de categorias na primeira execução do programa/rodada
        if self.contador_sorteios == 0 and (self.categorias_pool is None or len(self.categorias_pool) == 0):
            # Faz uma cópia da lista global com as repetições
            self.categorias_pool = list(self.categorias_pool_backup)
            # Comentário: esta cópia será consumida a cada sorteio até esvaziar

        participantes_disponiveis = [p for p in self.participantes 
                                     if p['id'] not in self.participantes_sorteados_ids]
        
        if not participantes_disponiveis:
            print("Todos sorteados! Reiniciando...")
            self.participantes_sorteados_ids.clear()
            participantes_disponiveis = self.participantes
        
        # Seleciona a categoria da vez a partir do pool e remove do pool
        # Se o pool estiver vazio (fim de rodada), reinicia a partir do backup
        if self.categorias_pool is None or len(self.categorias_pool) == 0:
            # Comentário: reinício do pool de categorias ao completar a rodada de 31 entradas
            self.categorias_pool = list(self.categorias_pool_backup)
        
        # Escolhe e consome uma categoria do pool
        categoria_escolhida = random.choice(self.categorias_pool)
        # Remove apenas uma ocorrência da categoria escolhida
        try:
            self.categorias_pool.remove(categoria_escolhida)
        except ValueError:
            # Comentário: improvável, mas se não encontrar, segue sem remover
            pass
        
        print(f"Categoria sorteada: {categoria_escolhida}")  # debug/output da categoria
        
        # Filtra participantes disponíveis pela categoria escolhida
        candidatos_categoria = [p for p in participantes_disponiveis if p.get('categoria') == categoria_escolhida]
        
        # Caso não haja candidatos para a categoria escolhida, tentamos recuperar a elegibilidade geral
        if not candidatos_categoria:
            # Se todos já foram sorteados anteriormente, limpamos para permitir nova rodada de pessoas
            if len(participantes_disponiveis) == 0:
                self.participantes_sorteados_ids.clear()
                participantes_disponiveis = list(self.participantes)
                candidatos_categoria = [p for p in participantes_disponiveis if p.get('categoria') == categoria_escolhida]
            
            # Se ainda assim não houver ninguém nessa categoria (ex.: categoria inexistente na planilha),
            # realizamos um fallback para qualquer participante disponível para não travar o sorteio.
            # Risco: isso pode quebrar a distribuição desejada. Sugestão: validar categorias com a planilha antes do evento.
            if not candidatos_categoria:
                # Fallback controlado: escolhe de todos disponíveis (mantido para robustez)
                candidatos_categoria = list(participantes_disponiveis)

        if not candidatos_categoria:
            # Se ainda assim não houver candidatos, retorna None (situação anômala)
            return None
        
        sorteado = random.choice(candidatos_categoria)
        self.participantes_sorteados_ids.add(sorteado['id'])
        
        return sorteado
    
    def iniciar_sorteio(self):
        """Inicia a animação do sorteio"""
        global audio_atual_index, tempo_inicio_audio, audio_em_execucao, audio_atual
        
        if not self.sorteando:
            self.sorteando = True
            self.velocidade_sorteio = 30
            self.tempo_sorteio = pygame.time.get_ticks()
            self.tempo_finalizacao = 0
            self.animacao_atual = 0
            
            # CORREÇÃO: Usa apenas UM áudio por sorteio, na sequência correta
            if audio_ativado and audios_sorteio:
                # Calcula qual áudio usar baseado no contador de sorteios
                audio_index = self.contador_sorteios % len(audios_sorteio)
                audio_atual_index = audio_index
                
                # Pega o áudio correspondente
                self.audio_sorteio_atual = audios_sorteio[audio_index]
                
                if self.audio_sorteio_atual:
                    # Para qualquer áudio que esteja tocando
                    mixer.stop()
                    
                    # Toca o áudio do sorteio atual
                    self.audio_sorteio_atual.play()
                    tempo_inicio_audio = pygame.time.get_ticks()
                    audio_em_execucao = True
                    audio_atual = self.audio_sorteio_atual
                    
                    print(f"🎵 Tocando áudio do sorteio #{self.contador_sorteios + 1}")
                else:
                    print(f"⚠ Áudio {audio_index + 1} não disponível!")
                    audio_em_execucao = False
            else:
                print(f"⚠ Áudio não disponível. Ativado: {audio_ativado}")
                audio_em_execucao = False
            
            print(f"▶ Sorteio #{self.contador_sorteios + 1} iniciado!")
    
    def finalizar_sorteio(self):
        """Finaliza o sorteio e salva no log"""
        global audio_em_execucao, audio_atual
        
        self.sorteando = False
        
        # Incrementa o contador de sorteios para próxima sequência de áudio
        self.contador_sorteios += 1
        
        # Sorteia o participante
        self.participante_sorteado = self.sortear_participante()
        
        if self.participante_sorteado:
            self.historico.append(self.participante_sorteado)
            salvar_log_sorteio(self.participante_sorteado)
            self.ultimo_sorteio = datetime.now().strftime("%H:%M:%S")
            
            # Para o áudio do sorteio que estava tocando
            audio_em_execucao = False
            mixer.stop()
            
            # Toca som de sucesso se disponível
            if audio_ativado and som_sucesso:
                som_sucesso.play()
                print("🎉 Áudio de sucesso!")
            
            print(f"✅ Sorteado: {self.participante_sorteado['nome']}")
    
    def atualizar(self):
        """Atualiza a lógica do sorteio e áudios"""
        global audio_atual_index, tempo_inicio_audio, audio_em_execucao, audio_atual
        
        tempo_atual = pygame.time.get_ticks()
        
        if self.sorteando:
            self.animacao_atual = (self.animacao_atual + 1) % 60
            
            # CORREÇÃO: Verifica se o áudio do sorteio já terminou (8 segundos)
            if audio_em_execucao and audio_ativado and audio_atual:
                tempo_decorrido = tempo_atual - tempo_inicio_audio
                
                # Se passaram 8 segundos, finaliza o sorteio automaticamente
                if tempo_decorrido > duracao_audio:  # 8000ms = 8 segundos
                    print("⏰ 8 segundos completos, finalizando sorteio...")
                    self.finalizar_sorteio()
                    return
            
            # Atualiza animação do sorteio
            if tempo_atual - self.tempo_sorteio > self.velocidade_sorteio:
                if self.participantes:
                    self.participante_sorteado = random.choice(self.participantes)
                    self.tempo_sorteio = tempo_atual
                    
                    # Aumenta velocidade gradualmente para durar ~8 segundos
                    if self.velocidade_sorteio < 500:  # Ajustado para 8 segundos
                        self.velocidade_sorteio += 10  # Ajustado para durar ~8s
                    else:
                        if self.tempo_finalizacao == 0:
                            self.tempo_finalizacao = tempo_atual
                        elif tempo_atual - self.tempo_finalizacao > 1000:
                            self.finalizar_sorteio()
    
    def criar_texto_responsivo(self, texto, max_largura):
        """Cria texto que se ajusta ao espaço disponível"""
        tamanho = 80
        
        while tamanho > 30:
            fonte_test = pygame.font.SysFont('arial', tamanho, bold=True)
            largura_texto = fonte_test.size(texto)[0]
            
            if largura_texto <= max_largura:
                return fonte_test.render(texto, True, COR_DESTAQUE)
            
            tamanho -= 5
        
        fonte_final = pygame.font.SysFont('arial', 30, bold=True)
        return fonte_final.render(texto, True, COR_DESTAQUE)
    
    def desenhar(self, tela):
        """Desenha todos os elementos na tela"""
        global largura_tela, altura_tela
        
        # Fundo
        tela.fill(COR_FUNDO)
        
        # Título
        titulo = fonte_media.render("SORTEADOR INRAD", True, COR_DESTAQUE)
        tela.blit(titulo, (largura_tela//2 - titulo.get_width()//2, 20))
        
        # Área do resultado
        area_altura = 400
        area_y = 80
        
        area_rect = pygame.Rect(50, area_y, largura_tela - 100, area_altura)
        pygame.draw.rect(tela, (30, 40, 60), area_rect, border_radius=15)
        pygame.draw.rect(tela, (60, 80, 120), area_rect, 3, border_radius=15)
        
        if self.sorteando:
            pygame.draw.rect(tela, (255, 255, 200, 50), 
                           area_rect.inflate(10, 10), border_radius=20, width=2)
        
        # Conteúdo da área
        if self.sorteando or self.participante_sorteado:
            if self.sorteando:
                titulo_texto = "SORTEANDO..."
                cor_titulo = COR_AMARELO
            else:
                titulo_texto = "PARABÉNS!"
                cor_titulo = COR_VERDE
            
            titulo_surf = fonte_media.render(titulo_texto, True, cor_titulo)
            tela.blit(titulo_surf, (largura_tela//2 - titulo_surf.get_width()//2, area_y + 30))
            
            if self.participante_sorteado:
                nome = self.participante_sorteado['nome']
                max_largura_nome = largura_tela - 200
                
                nome_surf = self.criar_texto_responsivo(nome, max_largura_nome)
                nome_rect = nome_surf.get_rect(center=(largura_tela//2, area_y + area_altura//2))
                tela.blit(nome_surf, nome_rect)

                # NOVO: Exibir a categoria logo abaixo do nome, centralizada e 80% menor que o texto do nome
                # Comentários (PT-BR):
                # - Ajustamos dinamicamente o tamanho com base no tamanho real renderizado do nome.
                # - Risco: em telas muito pequenas, o tamanho pode ficar pequeno demais; considerar limite mínimo.
                categoria = str(self.participante_sorteado.get('categoria', '')).strip()
                if categoria:
                    # Calcula tamanho da fonte da categoria como 20% do tamanho do nome (80% menor)
                    # Derivamos o tamanho aproximado do nome pelo menor lado do glyph box; fallback para 48.
                    try:
                        # Heurística: estimar tamanho base a partir da altura do surface do nome
                        tamanho_nome_px = nome_surf.get_height()
                        tamanho_categoria = max(12, int(tamanho_nome_px * 0.20))  # mínimo de 12px
                    except Exception:
                        tamanho_categoria = 12

                    fonte_categoria = pygame.font.SysFont('arial', tamanho_categoria, bold=False)
                    categoria_surf = fonte_categoria.render(categoria, True, COR_CATEGORIA)
                    categoria_rect = categoria_surf.get_rect(center=(largura_tela//2, nome_rect.bottom + 20))
                    tela.blit(categoria_surf, categoria_rect)
                
        else:
            mensagem = fonte_grande.render("PRONTO PARA SORTEAR!", True, (180, 200, 255))
            tela.blit(mensagem, (largura_tela//2 - mensagem.get_width()//2, area_y + area_altura//2 - 50))
            
            instrucao = fonte_normal.render("Clique em 'SORTEAR' para começar", True, (150, 180, 220))
            tela.blit(instrucao, (largura_tela//2 - instrucao.get_width()//2, area_y + area_altura//2 + 30))
        
        # BOTÕES (mais para cima, sem a barra)
        self.botao_sortear.desenhar(tela)
        
        # Estatísticas (mais para cima, sem a barra)
        stats_y = altura_tela - 120  # Subiu 40px
        stats_text = fonte_normal.render(
            f"Total: {len(self.participantes)} | "
            f"Sorteados: {len(self.historico)} | "
            f"Restantes: {len(self.participantes) - len(self.historico)}", 
            True, COR_TEXTO
        )
        tela.blit(stats_text, (largura_tela//2 - stats_text.get_width()//2, stats_y))
        
        # REMOVIDO: Informação do próximo áudio
        # REMOVIDO: Informação do último sorteio
        
        # RODAPÉ com informações do sistema
        rodape_altura = 40
        rodape_y = altura_tela - rodape_altura
        
        pygame.draw.rect(tela, COR_RODAPE, (0, rodape_y, largura_tela, rodape_altura))
        
        texto_rodape = fonte_rodape.render("Projetos & Produtos Digitais", True, COR_DESTAQUE)
        tela.blit(texto_rodape, (largura_tela//2 - texto_rodape.get_width()//2, rodape_y + 10))
        
        # Atalhos acima do rodapé
        atalhos_text = fonte_pequena.render(
            f"F11: Tela Cheia | F1: Áudio {'ON' if audio_ativado else 'OFF'} | ESC: Sair", 
            True, (150, 150, 150)
        )
        tela.blit(atalhos_text, (largura_tela//2 - atalhos_text.get_width()//2, rodape_y - 25))
        
        # Informação do arquivo na parte superior direita
        info_arquivo = fonte_pequena.render(
            f"Planilha: {os.path.basename(CAMINHO_EXCEL)}", 
            True, (180, 180, 200)
        )
        tela.blit(info_arquivo, (largura_tela - info_arquivo.get_width() - 20, 20))

# Cria o sorteador
sorteador = Sorteador()

# Loop principal
clock = pygame.time.Clock()
executando = True

# Tela de boas-vindas
print("\n" + "="*60)
print("🎉 SORTEADOR INRAD - PRONTO PARA USO!")
print("="*60)
print(f"📊 Participantes carregados: {len(sorteador.participantes)}")
print(f"📁 Planilha usada: {os.path.basename(CAMINHO_EXCEL)}")
print(f"📝 Log salvo em: {CAMINHO_LOG}")
print(f"🎵 Áudio carregados: {sum(1 for a in audios_sorteio if a is not None)}/13")
print(f"⏱️ Duração do áudio: {duracao_audio/1000} segundos por sorteio")
print("="*60 + "\n")

while executando:
    tempo_atual = pygame.time.get_ticks()
    mouse_pos = pygame.mouse.get_pos()
    
    for evento in pygame.event.get():
        if evento.type == pygame.QUIT:
            executando = False
            
        elif evento.type == pygame.VIDEORESIZE:
            if not tela_cheia:
                largura_tela, altura_tela = evento.w, evento.h
                tela = pygame.display.set_mode((largura_tela, altura_tela), pygame.RESIZABLE)
                sorteador.atualizar_botoes()
            
        elif evento.type == pygame.KEYDOWN:
            if evento.key == pygame.K_F11:
                tela_cheia = not tela_cheia
                if tela_cheia:
                    info = pygame.display.Info()
                    largura_tela, altura_tela = info.current_w, info.current_h
                    tela = pygame.display.set_mode((largura_tela, altura_tela), pygame.FULLSCREEN)
                else:
                    largura_tela, altura_tela = LARGURA, ALTURA
                    tela = pygame.display.set_mode((largura_tela, altura_tela), pygame.RESIZABLE)
                
                sorteador.atualizar_botoes()
                
            elif evento.key == pygame.K_F1:
                audio_ativado = not audio_ativado
                if not audio_ativado:
                    mixer.stop()
                print(f"🔊 Áudio: {'ATIVADO' if audio_ativado else 'DESATIVADO'}")
                
            elif evento.key == pygame.K_SPACE and not sorteador.sorteando:
                sorteador.iniciar_sorteio()
                
            elif evento.key == pygame.K_ESCAPE:
                if tela_cheia:
                    tela_cheia = False
                    largura_tela, altura_tela = LARGURA, ALTURA
                    tela = pygame.display.set_mode((largura_tela, altura_tela), pygame.RESIZABLE)
                    sorteador.atualizar_botoes()
                else:
                    executando = False
            
        elif evento.type == pygame.MOUSEBUTTONDOWN:
            if sorteador.botao_sortear.verificar_clique(mouse_pos) and not sorteador.sorteando:
                sorteador.iniciar_sorteio()
    
    # Atualiza hover dos botões
    sorteador.botao_sortear.verificar_hover(mouse_pos)
    
    # Atualiza lógica do sorteador
    sorteador.atualizar()
    
    # Desenha tudo
    sorteador.desenhar(tela)
    
    pygame.display.flip()
    clock.tick(60)

# Salva histórico final na área de trabalho
try:
    with open(CAMINHO_LOG, 'a', encoding='utf-8') as f:
        f.write(f"\n{'='*60}\n")
        f.write(f"Sessão finalizada: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write(f"Total de sorteados: {len(sorteador.historico)}\n")
        f.write(f"{'='*60}\n\n")
    print(f"📝 Log final salvo na área de trabalho: {CAMINHO_LOG}")
except Exception as e:
    print(f"✗ Erro ao salvar log final: {e}")

print("\n" + "="*60)
print("👋 SORTEADOR INRAD ENCERRADO")
print(f"🎯 Total de sorteados nesta sessão: {len(sorteador.historico)}")
print(f"🔢 Sequência de áudios usada: {sorteador.contador_sorteios} sorteios")
print("="*60 + "\n")

pygame.quit()
sys.exit()

## funcionando