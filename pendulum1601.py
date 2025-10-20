import init
from init import *

course_ID = "1601"
lab = "Pendulum_Dynamics"
excelbackup(txl,lab)

def f(θ0,ω0,ω):
  fig1 = plt.figure(1)

  def du_dx(U,x):
    return [U[1],-(ω**2)*np.sin(U[0])]

  def du_dxl(U,x):
    return [U[1],-(ω**2)*U[0]]

  U0 = [θ0,ω0]
  tmax = 10
  n = 1000
  step = 1/n
  xs = np.linspace(0,tmax,n)
  Us = odeint(du_dx,U0,xs)
  ys = Us[:,0]
  Usl = odeint(du_dxl,U0,xs)
  ysl = Usl[:,0]

  dθ_dt = np.diff(ys)/step
  dθ_dt_l = np.diff(ysl)/step

  plt.plot(xs, ysl,label='Linear Solution')
  plt.plot(xs, ys,label='Nonlinear Solution')
  plt.xlim(0,tmax)
  plt.ylim()
  plt.xlabel("t [s]")
  plt.ylabel("θ [rad]")
  plt.title("θ(t) Plot")
  plt.grid(color='silver', linestyle='--', linewidth=1)
  plt.legend(bbox_to_anchor=(-0.5,1), loc="upper left")
  plt.show()
  fig1.savefig('Figure2.png', bbox_inches = "tight", dpi=600)

  fig2 = plt.figure(1)
  plt.plot(dθ_dt_l, ysl[:-1],label='Linear Solution')
  plt.plot(dθ_dt, ys[:-1],label='Nonlinear Solution')
  plt.xlim()
  plt.ylim()
  plt.xlabel("dθ/dt [rad/s]")
  plt.ylabel("θ [rad]")
  plt.title("Phase Space")
  plt.grid(color='silver', linestyle='--', linewidth=1)
  plt.legend(bbox_to_anchor=(-0.5,1), loc="upper left")
  plt.show()
  fig2.savefig('Figure3.png', bbox_inches = "tight", dpi=600)


#plots the data and sliders
def labplot():
    return interact(f, θ0=FloatSlider(min=0, max=np.pi, step=0.01, value=1.8, description='θ_(t=0):'), ω0=FloatSlider(min=0, max=10, step=0.01,value=2.8, description='dθ/dt_(t=0):'), ω=FloatSlider(min=0, max=10, step=0.01,value=2.3, description='ω:'));

def img(filename):
    from IPython.display import Image, display
    return Image(filename, width=500)

def savepdf():
    import datetime
    import openpyxl
    
    from openpyxl import load_workbook
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, inch, portrait
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, PageBreak, Spacer
    from reportlab.lib.styles import getSampleStyleSheet

    answers = [Q_1.value,Q_2.value,Q_3.value,Q_4.value,Q_5.value,Q_6.value,Q_7.value,Q_8.value,Q_9.value,Q_10.value,Q_11.value,Q_12.value,Q_13.value,Q_14.value,Q_15.value,Q_16.value,Q_17.value,Q_18.value]
    workbook = load_workbook(filename=lab+txl+".xlsx")
    sheet = workbook.active
    for i in range(len(answers)):
        sheet["B"+str(i+1)]=answers[i]
    workbook.save(filename=lab+txl+".xlsx")

    Course_Number = "1601Q"
    Lab_Number = "QL08"
    file_timestamp = datetime.datetime.today().strftime('%Y-%m')
    Section_Number = section.value[0:3]

    doc = SimpleDocTemplate(str(file_timestamp)+"_"+str(Course_Number)+
                        "_"+str(Section_Number)+"_"+str(Lab_Number)+
                        "_"+str(group.value)+".pdf",
                        pagesize=letter, rightMargin=50,
                        leftMargin=50, topMargin=50,bottomMargin=50)
    doc.pagesize = portrait(letter)
    elements = []

    timestamp = datetime.datetime.now().isoformat()

    fig1 = Image("Figure1.png", width=200, height = 132)
    fig1.hAlign = 'LEFT'

    fig2 = Image("Trial_theta.png", width=200, height = 132)
    fig2.hAlign = 'LEFT'

    fig3 = Image("Trial_phase.png", width=200, height = 132)
    fig3.hAlign = 'LEFT'

    data = [
    ["timestamp:", str(timestamp)],
    ["Course_Number:", str(Course_Number)],
    ["Section_Number:", str(Section_Number)],
    ["Lab_Number:", str(Lab_Number)],
    ["Group_Number:", str(group.value)],
    ["Group_Members:", str(names.value)],
    ["Q1:", Q_1.value],
    ["Q2:", Q_2.value],
    ["Q3:", Q_3.value],
    ["Q4:", Q_4.value],
    ["Q5:", Q_5.value],
    ["Q6:", Q_6.value],
    ["Q7:", Q_7.value],
    ["Q10:", Q_10.value],
    ["Q11:", Q_11.value],
    ["Q12:", Q_12.value],
    ["Q13:", Q_13.value],
    ["Q14:", Q_14.value],
    ["Q15:", Q_15.value],
    ["Q16:", Q_16.value],
    ]

    figdata = [
    ["Experimental Results", "Trial θ(t) Plot"],
    [[fig1], [fig2]],
    ["Trial Phase Space Plot", ""],
    [[fig3], ""],
    ]


    #TODO: Get this line right instead of just copying it from the docs
    style = TableStyle([('ALIGN',(1,1),(-2,-2),'RIGHT'),
                       ('TEXTCOLOR',(1,1),(-2,-2),colors.black),
                       ('VALIGN',(0,0),(0,-1),'TOP'),
                       ('TEXTCOLOR',(0,0),(0,-1),colors.black),
                       ('ALIGN',(0,-1),(-1,-1),'CENTER'),
                       ('VALIGN',(0,-1),(-1,-1),'MIDDLE'),
                       ('TEXTCOLOR',(0,-1),(-1,-1),colors.black),
                       ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                       ('BOX', (0,0), (-1,-1), 0.25, colors.black),
                       ])

    #Configure style and word wrap
    s = getSampleStyleSheet()
    s = s["BodyText"]
    s.wordWrap = 'CJK'
    data2 = [[Paragraph(cell, s) for cell in row] for row in data]
    t=Table(data2,colWidths=(100,None))
    t2=Table(figdata)
    t.setStyle(style)
    t2.setStyle(style)

    #Send the data and build the file
    elements.append(t)
    elements.append(PageBreak())
    elements.append(t2)
    doc.build(elements)
