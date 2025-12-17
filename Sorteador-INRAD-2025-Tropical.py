import pygame
import random
import sys
import pandas as pd
from datetime import datetime
import os
import math
from pygame import mixer

# Inicializa o Pygame e o mixer de áudio
pygame.init()
mixer.init()

# Configurações da tela
LARGURA, ALTURA = 1200, 800
tela = pygame.display.set_mode((LARGURA, ALTURA), pygame.RESIZABLE)
pygame.display.set_caption("CONFRATERNIZAÇÃO INRAD 2025")

# Variáveis globais
tela_cheia = False
audio_ativado = True
largura_tela, altura_tela = LARGURA, ALTURA

# Obtém o caminho da área de trabalho do usuário
CAMINHO_DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")

# Caminhos dos arquivos na área de trabalho
CAMINHO_EXCEL = os.path.join(CAMINHO_DESKTOP, "sorteio.xlsx")
CAMINHO_LOG = os.path.join(CAMINHO_DESKTOP, "log_sorteios_tropical.txt")

print(f"Procurando planilha em: {CAMINHO_EXCEL}")
print(f"Log será salvo em: {CAMINHO_LOG}")

# CORES TROPICAIS VIBRANTES
COR_FUNDO = (20, 120, 150)  # Azul tropical
COR_TEXTO = (255, 255, 255)
COR_DESTAQUE = (255, 215, 0)  # Dourado
COR_DESTAQUE2 = (255, 100, 0)  # Laranja
COR_CATEGORIA = (0, 200, 180)  # Turquesa
COR_BOTAO = (255, 105, 180)  # Rosa tropical
COR_BOTAO_HOVER = (255, 140, 200)  # Rosa claro
COR_BOTAO_SORTEIO = (255, 69, 0)  # Vermelho laranja
COR_BOTAO_SORTEIO_HOVER = (255, 99, 30)
COR_VERDE = (50, 205, 50)  # Verde limão
COR_AMARELO = (255, 215, 0)  # Amarelo dourado
COR_AZUL = (30, 144, 255)  # Azul royal
COR_RODAPE = (40, 100, 80)  # Verde tropical escuro

# Gradiente de fundo tropical simplificado
def desenhar_fundo_tropical(superficie, largura, altura):
    """Desenha um fundo tropical simplificado"""
    # Gradiente azul
    for i in range(altura):
        # Interpolação entre azul claro e escuro
        r = int(20 + (10-20) * i/altura)
        g = int(120 + (80-120) * i/altura)
        b = int(150 + (130-150) * i/altura)
        pygame.draw.line(superficie, (r, g, b), (0, i), (largura, i))
    
    # Sol tropical
    pygame.draw.circle(superficie, (255, 215, 0), (largura-100, 100), 60)
    
    # Brilho do sol
    for i in range(1, 4):
        raio = 60 + i * 10
        superficie_brilho = pygame.Surface((raio*2, raio*2), pygame.SRCALPHA)
        alpha = 100 - i * 25
        pygame.draw.circle(superficie_brilho, (255, 255, 150, alpha), (raio, raio), raio)
        superficie.blit(superficie_brilho, (largura-100-raio, 100-raio))

# Fontes
try:
    fonte_grande = pygame.font.SysFont('comicsansms', 52, bold=True)
    fonte_media = pygame.font.SysFont('comicsansms', 36, bold=True)
    fonte_normal = pygame.font.SysFont('comicsansms', 28)
    fonte_pequena = pygame.font.SysFont('comicsansms', 20)
    fonte_rodape = pygame.font.SysFont('comicsansms', 24, bold=True)
except:
    # Fallback para fontes padrão
    fonte_grande = pygame.font.SysFont(None, 52, bold=True)
    fonte_media = pygame.font.SysFont(None, 36, bold=True)
    fonte_normal = pygame.font.SysFont(None, 28)
    fonte_pequena = pygame.font.SysFont(None, 20)
    fonte_rodape = pygame.font.SysFont(None, 24, bold=True)

# Carrega dados do Excel
def carregar_dados_excel():
    try:
        if not os.path.exists(CAMINHO_EXCEL):
            print("⚠ Planilha não encontrada na área de trabalho!")
            print(f"⚠ Procurei em: {CAMINHO_EXCEL}")
            
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
    colunas_df = [str(col).upper().strip() for col in df.columns]
    
    print(f"Colunas encontradas: {colunas_df}")
    
    categoria_idx = None
    id_idx = None
    nome_idx = None
    
    for i, col in enumerate(colunas_df):
        col_clean = col.replace(" ", "").replace("_", "").replace("-", "")
        
        if any(keyword in col_clean for keyword in ['CATEGORIA', 'SETOR', 'DEPARTAMENTO', 'GRUPO', 'TIPO']):
            categoria_idx = i
        elif any(keyword in col_clean for keyword in ['ID', 'MATRICULA', 'CODIGO', 'NÚMERO', 'NUMERO', 'COD']):
            id_idx = i
        elif any(keyword in col_clean for keyword in ['NOME', 'NOMES', 'PARTICIPANTE', 'FUNCIONARIO', 'PESSOA']):
            nome_idx = i
    
    if categoria_idx is None:
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
            
            cat = cat.strip()
            id_val = str(id_val).strip()
            nome = nome.strip()
            
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
    
    exemplo_excel = os.path.join(CAMINHO_DESKTOP, "EXEMPLO-SORTEIO-TROPICAL.xlsx")
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
        
        if not os.path.exists(CAMINHO_DESKTOP):
            os.makedirs(CAMINHO_DESKTOP, exist_ok=True)
            print(f"Criada pasta da área de trabalho: {CAMINHO_DESKTOP}")
        
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
        self.sombra = pygame.Rect(x+3, y+3, largura, altura)
        
    def calcular_cor_hover(self, cor):
        return tuple(min(c + 40, 255) for c in cor)
        
    def desenhar(self, superficie):
        # Desenha sombra
        pygame.draw.rect(superficie, (0, 0, 0, 100), self.sombra, border_radius=12)
        
        # Desenha botão com borda decorativa
        pygame.draw.rect(superficie, self.cor, self.rect, border_radius=10)
        
        # Efeito de brilho no topo
        pygame.draw.line(superficie, (255, 255, 255, 150), 
                        (self.rect.left+2, self.rect.top+2), 
                        (self.rect.right-2, self.rect.top+2), 2)
        
        # Borda decorativa
        pygame.draw.rect(superficie, (255, 255, 255), self.rect, 3, border_radius=10)
        
        # Texto com sombra
        try:
            fonte_botao = pygame.font.SysFont('comicsansms', 24, bold=True)
        except:
            fonte_botao = pygame.font.SysFont(None, 24, bold=True)
            
        texto_surf = fonte_botao.render(self.texto, True, COR_TEXTO)
        texto_sombra = fonte_botao.render(self.texto, True, (0, 0, 0))
        
        texto_rect = texto_surf.get_rect(center=self.rect.center)
        sombra_rect = texto_sombra.get_rect(center=(self.rect.centerx+2, self.rect.centery+2))
        
        superficie.blit(texto_sombra, sombra_rect)
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
        self.tempo_inicio = pygame.time.get_ticks()
        
        self.contador_sorteios = 0
        self.audio_sorteio_atual = None
        
        # Elementos decorativos animados (apenas folhas)
        self.folhas_animacao = []
        self.inicializar_folhas()
        
        self.atualizar_botoes()
    
    def inicializar_folhas(self):
        """Inicializa folhas caindo para animação"""
        for _ in range(15):  # Menos folhas
            self.folhas_animacao.append({
                'x': random.randint(0, largura_tela),
                'y': random.randint(-100, 0),
                'vel_x': random.uniform(-0.5, 0.5),
                'vel_y': random.uniform(1, 2),
                'tamanho': random.randint(8, 20),
                'cor': random.choice([(255, 215, 0), (255, 140, 0), (50, 205, 50)]),  # Cores tropicais
                'rotacao': random.uniform(0, 360),
            })
    
    def atualizar_animacao_folhas(self):
        """Atualiza animação das folhas caindo"""
        for folha in self.folhas_animacao:
            folha['x'] += folha['vel_x']
            folha['y'] += folha['vel_y']
            folha['rotacao'] += folha['vel_x'] * 2
            
            # Reinicia folha se sair da tela
            if folha['y'] > altura_tela:
                folha['x'] = random.randint(0, largura_tela)
                folha['y'] = random.randint(-100, 0)
                folha['vel_y'] = random.uniform(1, 2)
    
    def desenhar_folhas(self, superficie):
        """Desenha folhas caindo"""
        for folha in self.folhas_animacao:
            # Desenha folha como um pequeno triângulo
            pontos = []
            for i in range(3):
                angulo = folha['rotacao'] + i * 120
                rad_angle = math.radians(angulo)
                px = folha['x'] + folha['tamanho'] * 0.5 * math.cos(rad_angle)
                py = folha['y'] + folha['tamanho'] * 0.3 * math.sin(rad_angle)
                pontos.append((px, py))
            
            pygame.draw.polygon(superficie, folha['cor'], pontos)
    
    def atualizar_botoes(self):
        global largura_tela, altura_tela
        
        btn_largura = 280
        btn_altura = 60
        espacamento = 30
        
        pos_y = altura_tela - 180
        
        total_largura = (btn_largura * 2) + espacamento
        inicio_x = (largura_tela - total_largura) // 2
        
        self.botao_sortear = Botao(inicio_x, pos_y, btn_largura, btn_altura, 
                                   "INICIAR SORTEIO", COR_BOTAO_SORTEIO)
        self.botao_novo = Botao(inicio_x + btn_largura + espacamento, pos_y, 
                               btn_largura, btn_altura, "NOVO SORTEIO", COR_AZUL)
    
    def sortear_participante(self):
        if not self.participantes:
            return None
        
        participantes_disponiveis = [p for p in self.participantes 
                                    if p['id'] not in self.participantes_sorteados_ids]
        
        if not participantes_disponiveis:
            print("Todos sorteados! Reiniciando...")
            self.participantes_sorteados_ids.clear()
            participantes_disponiveis = self.participantes
        
        sorteado = random.choice(participantes_disponiveis)
        self.participantes_sorteados_ids.add(sorteado['id'])
        
        return sorteado
    
    def iniciar_sorteio(self):
        if not self.sorteando:
            self.sorteando = True
            self.velocidade_sorteio = 30
            self.tempo_sorteio = pygame.time.get_ticks()
            self.tempo_finalizacao = 0
            self.animacao_atual = 0
            
            print(f"▶ Sorteio #{self.contador_sorteios + 1} iniciado!")
    
    def finalizar_sorteio(self):
        self.sorteando = False
        self.contador_sorteios += 1
        
        self.participante_sorteado = self.sortear_participante()
        
        if self.participante_sorteado:
            self.historico.append(self.participante_sorteado)
            salvar_log_sorteio(self.participante_sorteado)
            self.ultimo_sorteio = datetime.now().strftime("%H:%M:%S")
            
            print(f"✅ Sorteado: {self.participante_sorteado['nome']}")
    
    def atualizar(self):
        tempo_atual = pygame.time.get_ticks()
        
        # Atualiza animação das folhas
        self.atualizar_animacao_folhas()
        
        if self.sorteando:
            self.animacao_atual = (self.animacao_atual + 2) % 60
            
            # Atualiza animação do sorteio
            if tempo_atual - self.tempo_sorteio > self.velocidade_sorteio:
                if self.participantes:
                    self.participante_sorteado = random.choice(self.participantes)
                    self.tempo_sorteio = tempo_atual
                    
                    if self.velocidade_sorteio < 500:
                        self.velocidade_sorteio += 10
                    else:
                        if self.tempo_finalizacao == 0:
                            self.tempo_finalizacao = tempo_atual
                        elif tempo_atual - self.tempo_finalizacao > 1000:
                            self.finalizar_sorteio()
    
    def criar_texto_responsivo(self, texto, max_largura, cor=(0, 0, 0)):
        """Cria texto que se ajusta ao espaço disponível"""
        tamanho = 80
        
        while tamanho > 30:
            try:
                fonte_test = pygame.font.SysFont('comicsansms', tamanho, bold=True)
            except:
                fonte_test = pygame.font.SysFont(None, tamanho, bold=True)
                
            largura_texto = fonte_test.size(texto)[0]
            
            if largura_texto <= max_largura:
                return fonte_test.render(texto, True, cor)
            
            tamanho -= 5
        
        try:
            fonte_final = pygame.font.SysFont('comicsansms', 30, bold=True)
        except:
            fonte_final = pygame.font.SysFont(None, 30, bold=True)
            
        return fonte_final.render(texto, True, cor)
    
    def desenhar(self, tela):
        global largura_tela, altura_tela
        
        tempo = pygame.time.get_ticks()
        
        # Fundo tropical simplificado
        desenhar_fundo_tropical(tela, largura_tela, altura_tela)
        
        # Folhas caindo
        self.desenhar_folhas(tela)
        
        # Título com efeito
        titulo_texto = "CONFRATERNIZAÇÃO INRAD"
        titulo = fonte_media.render(titulo_texto, True, COR_DESTAQUE)
        
        # Sombra do título
        titulo_sombra = fonte_media.render(titulo_texto, True, (0, 0, 0))
        tela.blit(titulo_sombra, (largura_tela//2 - titulo.get_width()//2 + 3, 23))
        tela.blit(titulo, (largura_tela//2 - titulo.get_width()//2, 20))
        
        # Subtítulo
        subtitulo_text = "2025"
        subtitulo = fonte_pequena.render(subtitulo_text, True, (255, 255, 200))
        tela.blit(subtitulo, (largura_tela//2 - subtitulo.get_width()//2, 65))
        
        # Área do resultado - Estilo limpo e moderno
        area_altura = 400
        area_y = 100
        
        # Área principal (fundo branco limpo)
        area_rect = pygame.Rect(50, area_y, largura_tela - 100, area_altura)
        
        # Fundo branco limpo
        pygame.draw.rect(tela, (173, 173, 137), area_rect, border_radius=15)
        
        # Borda decorativa simples
        pygame.draw.rect(tela, COR_DESTAQUE, area_rect, 3, border_radius=15)
        
        # Efeito de luz durante sorteio (sutil)
        if self.sorteando:
            for i in range(3):
                alpha = 100 - i * 30
                superficie_luz = pygame.Surface((area_rect.width + i*15, area_rect.height + i*15), pygame.SRCALPHA)
                pygame.draw.rect(superficie_luz, (255, 255, 200, alpha), 
                               (0, 0, superficie_luz.get_width(), superficie_luz.get_height()), 
                               border_radius=15 + i*2)
                tela.blit(superficie_luz, (area_rect.x - i*7, area_rect.y - i*7))
        
        # Conteúdo da área
        if self.sorteando or self.participante_sorteado:
            if self.sorteando:
                titulo_texto = "SORTEANDO..."
                cor_titulo = COR_AMARELO
                
                # Efeito de piscar durante sorteio
                if tempo % 500 < 250:
                    cor_titulo = COR_DESTAQUE2
            else:
                titulo_texto = "PARABÉNS!"
                cor_titulo = COR_VERDE
            
            titulo_surf = fonte_media.render(titulo_texto, True, cor_titulo)
            titulo_sombra = fonte_media.render(titulo_texto, True, (0, 0, 0))
            
            tela.blit(titulo_sombra, (largura_tela//2 - titulo_surf.get_width()//2 + 2, area_y + 32))
            tela.blit(titulo_surf, (largura_tela//2 - titulo_surf.get_width()//2, area_y + 30))
            
            # Área para mostrar o nome
            nome_area_y = area_y + 100
            nome_area_altura = 200
            
            if self.participante_sorteado:
                nome = self.participante_sorteado['nome']
            else:
                # Durante sorteio, mostra nome aleatório
                if self.participantes:
                    nome = random.choice(self.participantes)['nome']
                else:
                    nome = "Carregando..."
            
            max_largura_nome = largura_tela - 200
            
            # NOME em PRETO durante sorteio, dourado quando sorteado
            if self.sorteando:
                cor_nome = (0, 0, 0)  # PRETO durante sorteio
            else:
                cor_nome = COR_DESTAQUE  # Dourado quando sorteado
            
            nome_surf = self.criar_texto_responsivo(nome, max_largura_nome, cor_nome)
            nome_rect = nome_surf.get_rect(center=(largura_tela//2, area_y + area_altura//2))
            tela.blit(nome_surf, nome_rect)
            
            # Informações adicionais apenas quando sorteado
            if not self.sorteando and self.participante_sorteado:
                info_y = area_y + area_altura//2 + 70
                info_texto = f"Categoria: {self.participante_sorteado['categoria']}  |  ID: {self.participante_sorteado['id']}"
                info_surf = fonte_normal.render(info_texto, True, COR_CATEGORIA)
                tela.blit(info_surf, (largura_tela//2 - info_surf.get_width()//2, info_y))
                
        else:
            # Tela inicial - pronto para sortear
            mensagem = fonte_grande.render("PRONTO PARA SORTEAR!", True, COR_DESTAQUE)
            sombra_msg = fonte_grande.render("PRONTO PARA SORTEAR!", True, (0, 0, 0))
            
            tela.blit(sombra_msg, (largura_tela//2 - mensagem.get_width()//2 + 2, area_y + area_altura//2 - 52))
            tela.blit(mensagem, (largura_tela//2 - mensagem.get_width()//2, area_y + area_altura//2 - 50))
            
            instrucao = fonte_normal.render("Clique no botão para começar o sorteio!", True, (100, 100, 120))
            tela.blit(instrucao, (largura_tela//2 - instrucao.get_width()//2, area_y + area_altura//2 + 30))
        
        # BOTÕES
        self.botao_sortear.desenhar(tela)
        self.botao_novo.desenhar(tela)
        
        # Estatísticas
        stats_y = altura_tela - 120
        stats_text = fonte_normal.render(
            f"Total: {len(self.participantes)} | "
            f"Sorteados: {len(self.historico)} | "
            f"Restantes: {len(self.participantes) - len(self.historico)}", 
            True, COR_TEXTO
        )
        stats_sombra = fonte_normal.render(
            f"Total: {len(self.participantes)} | "
            f"Sorteados: {len(self.historico)} | "
            f"Restantes: {len(self.participantes) - len(self.historico)}", 
            True, (0, 0, 0)
        )
        tela.blit(stats_sombra, (largura_tela//2 - stats_text.get_width()//2 + 2, stats_y + 2))
        tela.blit(stats_text, (largura_tela//2 - stats_text.get_width()//2, stats_y))
        
        # RODAPÉ
        rodape_altura = 40
        rodape_y = altura_tela - rodape_altura
        
        pygame.draw.rect(tela, COR_RODAPE, (0, rodape_y, largura_tela, rodape_altura))
        
        # Padrão decorativo no rodapé (simples)
        for i in range(0, largura_tela, 40):
            pygame.draw.circle(tela, COR_DESTAQUE, (i, rodape_y + rodape_altura//2), 3)
        
        texto_rodape = fonte_rodape.render("Projetos & Produtos Digitais", True, COR_DESTAQUE)
        tela.blit(texto_rodape, (largura_tela//2 - texto_rodape.get_width()//2, rodape_y + 8))
        
        # Atalhos
        atalhos_text = fonte_pequena.render(
            f"F11: Tela Cheia | F1: Áudio {'ON' if audio_ativado else 'OFF'} | 🚪 ESC: Sair", 
            True, (200, 255, 200)
        )
        tela.blit(atalhos_text, (largura_tela//2 - atalhos_text.get_width()//2, rodape_y - 30))
        
        # Informação do arquivo
        info_arquivo = fonte_pequena.render(
            f"Planilha: {os.path.basename(CAMINHO_EXCEL)}", 
            True, (255, 255, 200)
        )
        tela.blit(info_arquivo, (largura_tela - info_arquivo.get_width() - 20, 20))

# Cria o sorteador
sorteador = Sorteador()

# Loop principal
clock = pygame.time.Clock()
executando = True

print("\n" + "="*60)
print("SORTEADOR TROPICAL - PRONTO PARA USO!")
print("="*60)
print(f"Participantes carregados: {len(sorteador.participantes)}")
print(f"Planilha usada: {os.path.basename(CAMINHO_EXCEL)}")
print(f"Log salvo em: {CAMINHO_LOG}")
print(f"Interface limpa e moderna")
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
                print(f"Áudio: {'ATIVADO' if audio_ativado else 'DESATIVADO'}")
                
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
            
            elif sorteador.botao_novo.verificar_clique(mouse_pos):
                sorteador.sorteando = False
                sorteador.participante_sorteado = None
                print("Novo sorteio pronto!")
    
    # Atualiza hover dos botões
    sorteador.botao_sortear.verificar_hover(mouse_pos)
    sorteador.botao_novo.verificar_hover(mouse_pos)
    
    # Atualiza lógica do sorteador
    sorteador.atualizar()
    
    # Desenha tudo
    sorteador.desenhar(tela)
    
    pygame.display.flip()
    clock.tick(60)

# Salva histórico final
try:
    with open(CAMINHO_LOG, 'a', encoding='utf-8') as f:
        f.write(f"\n{'='*60}\n")
        f.write(f"Sessão Finalizada: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write(f"Total de sorteados: {len(sorteador.historico)}\n")
        f.write(f"{'='*60}\n\n")
    print(f"Log final salvo: {CAMINHO_LOG}")
except Exception as e:
    print(f"✗ Erro ao salvar log final: {e}")

print("\n" + "="*60)
print("SORTEADOR ENCERRADO")
print(f"Total de sorteados: {len(sorteador.historico)}")
print(f"Até a próxima!")
print("="*60 + "\n")

pygame.quit()
sys.exit()