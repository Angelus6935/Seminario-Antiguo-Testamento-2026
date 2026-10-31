from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()

# Colores
NEGRO = RGBColor(18, 18, 18)
CYAN = RGBColor(0, 240, 255)
ROJO = RGBColor(255, 0, 85)
BLANCO = RGBColor(255, 255, 255)


def slide_base():
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = NEGRO
    return slide


def add_text(slide, text, left, top, size=32, color=BLANCO, bold=True):
    box = slide.shapes.add_textbox(left, top, Inches(9), Inches(1.5))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


# 1 Portada
s = slide_base()
add_text(s, "¿CÓMO SOBREVIVIR AL FUTURO?", Inches(1), Inches(2), 44, CYAN)
add_text(s, "Ejercer la fe en Jesucristo para desarrollar la autosuficiencia", Inches(1), Inches(3.5), 24, BLANCO)

# 2 Under Pressure
s = slide_base()
add_text(s, "Pressure pushing down on me, pressing down on you...", Inches(1), Inches(2), 32, ROJO)
add_text(s, "Presiones reales\nColegio\nFuturo\nRedes\nAnsiedad\nLa autosuficiencia evita quebrarnos", Inches(1), Inches(3.5), 22)

# 3 José esfuerzo
s = slide_base()
add_text(s, "Génesis 39:2-3", Inches(1), Inches(1), 26, CYAN)
add_text(s, "Dios prosperaba lo que José hacía con su mano\nTenía 17 años\nEstaba solo\nEl esfuerzo activa la bendición", Inches(1), Inches(2.5), 22)

# 4 Cárcel resiliencia
s = slide_base()
add_text(s, "Génesis 39:21", Inches(1), Inches(1), 26, CYAN)
add_text(s, "Inocente pero preso\nMantuvo actitud\nAyudó a otros\nOrganizó la cárcel\nResiliencia emocional real", Inches(1), Inches(2.5), 22)

# 5 Fundamento
s = slide_base()
add_text(s, "Fundamento de la Autosuficiencia", Inches(1), Inches(1), 28, CYAN)
add_text(s, "Jesucristo es el fundamento\nSu gracia habilita\nNuestro esfuerzo responde\nResultado: autosuficiencia integral", Inches(1), Inches(2.5), 22)

# 6 Anime
s = slide_base()
add_text(s, "Adaptación y Estrategia", Inches(1), Inches(1), 28, CYAN)
add_text(s, "Isagi analiza errores\nSe adapta\nUsa inteligencia\nJosé hizo lo mismo\nDios bendice el estudio y la mente", Inches(1), Inches(2.5), 22)

# 7 127 horas
s = slide_base()
add_text(s, "Decisiones difíciles", Inches(1), Inches(1), 28, ROJO)
add_text(s, "Aron atrapado\nDecisión extrema para vivir\nDar el máximo esfuerzo\nCortar lo que frena el progreso", Inches(1), Inches(2.5), 22)

# 8 Mineros
s = slide_base()
add_text(s, "Organización en la oscuridad", Inches(1), Inches(1), 28, CYAN)
add_text(s, "33 mineros\nRacionaron comida\nCuidaron mente\nOraron\nSu esfuerzo los sostuvo\nDependían del rescate", Inches(1), Inches(2.5), 22)

# 9 Dinámica
s = slide_base()
add_text(s, "Simulador de Crisis", Inches(1), Inches(1), 28, ROJO)
add_text(s, "Crisis emocional\nCrisis económica\nCrisis espiritual\nCuál es nuestro esfuerzo\nCómo actúa Cristo", Inches(1), Inches(2.5), 22)

# 10 Áreas
s = slide_base()
add_text(s, "5 Áreas de Autosuficiencia", Inches(1), Inches(1), 28, CYAN)
add_text(s, "Evangelio\nEducación\nSalud física\nSalud emocional\nFinanzas\nDónde aplicarás esta semana", Inches(1), Inches(2.5), 22)

# 11 Conclusión
s = slide_base()
add_text(s, "De la cárcel al palacio", Inches(1), Inches(1), 28, CYAN)
add_text(s, "Pon tu máximo esfuerzo\nPermite que el poder del Salvador complete el milagro", Inches(1), Inches(2.5), 26)

prs.save("Seminario_Autosuficiencia.pptx")
print("Listo: Seminario_Autosuficiencia.pptx")
