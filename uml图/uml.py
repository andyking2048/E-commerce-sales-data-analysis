import plantuml
import os

def generate_uml_diagrams():
    """
    Generates Use Case, Activity, and Class diagrams for the
    Excel Data Analysis Tool.
    """
    # URL for the PlantUML server
    PLANTUML_URL = "http://www.plantuml.com/plantuml"

    # Create a PlantUML object
    p = plantuml.PlantUML(url=f"{PLANTUML_URL}/img/")

    # --- Use Case Diagram ---
    use_case_diagram = """
    @startuml
    left to right direction
    actor "Data Analyst" as user
    rectangle "Excel Data Analysis Tool" {
        usecase "Import Excel Files" as UC1
        usecase "View Data" as UC2
        usecase "Extract Columns" as UC3
        usecase "Filter Data" as UC4
        usecase "Merge Files" as UC5
        usecase "Generate Statistics" as UC6
        usecase "Create Chart" as UC7
        usecase "Export Results" as UC8
    }

    user --> UC1
    user --> UC2
    user --> UC3
    user --> UC4
    user --> UC5
    user --> UC6
    user --> UC7
    user --> UC8

    @enduml
    """

    # --- Activity Diagram ---
    activity_diagram = """
    @startuml
    title Activity Diagram: Data Analysis Workflow

    start
    :Select Directory with Excel Files;
    :Import Files;
    :Select a File to View;
    fork
        :Extract Specific Columns;
    fork again
        :Filter Data based on Criteria;
    fork again
        :Merge Multiple Files;
    end fork
    :Generate Statistical Summary;
    :Create Bar Chart for Visualization;
    :Save Results to a New Excel File;
    stop

    @enduml
    """

    # --- Class Diagram ---
    class_diagram = """
    @startuml
    class Ui_MainWindow {
        - centralwidget: QWidget
        - list1: QListView
        - text1: QTextEdit
        - viewButton: QPushButton
        - rButton1: QRadioButton
        - rButton2: QRadioButton
        - textEdit: QTextEdit
        - menubar: QMenuBar
        - statusbar: QStatusBar
        - toolBar: QToolBar
        - button1: QAction
        - button2: QAction
        - button3: QAction
        - button4: QAction
        - button5: QAction
        - button6: QAction
        - button7: QAction
        + setupUi(MainWindow): void
        + retranslateUi(MainWindow): void
        + click1(): void
        + clicked(qModelIndex): void
        + click2(): void
        + click3(): void
        + click4(): void
        + click5(): void
        + click6(): void
        + viewButton_click(): void
    }

    class pandas {
        + read_excel()
        + DataFrame()
        + concat()
        + groupby()
        + to_excel()
    }

    class matplotlib {
        + plot()
        + show()
        + figure()
    }

    Ui_MainWindow --|> QMainWindow
    Ui_MainWindow ..> pandas : uses
    Ui_MainWindow ..> matplotlib : uses

    @enduml
    """

    # Generate and save the diagrams
    try:
        if not os.path.exists("uml_diagrams"):
            os.makedirs("uml_diagrams")

        with open("uml_diagrams/use_case.puml", "w") as f:
            f.write(use_case_diagram)
        p.processes_file("uml_diagrams/use_case.puml")

        with open("uml_diagrams/activity.puml", "w") as f:
            f.write(activity_diagram)
        p.processes_file("uml_diagrams/activity.puml")

        with open("uml_diagrams/class.puml", "w") as f:
            f.write(class_diagram)
        p.processes_file("uml_diagrams/class.puml")

        print("UML diagrams generated successfully in the 'uml_diagrams' directory.")

    except Exception as e:
        print(f"An error occurred during UML diagram generation: {e}")

if __name__ == "__main__":
    generate_uml_diagrams()