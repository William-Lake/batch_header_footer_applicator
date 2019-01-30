import PySimpleGUI as sg

class PrimaryUI(sg.Window):

    DEFAULT_UPDATE_TEXT = 'Results....'

    def __init__(self):

        super().__init__('Batch Word Header Footer Replacement')

        self.__create_layout()
        
    def __create_layout(self):

        self.__txt_file_path_doc = sg.InputText(default_text = 'C:\Path\To\File_With_Target_Word_Doc_Paths.txt', do_not_clear = True)

        btn_browse_file_path_doc = sg.FileBrowse(file_types=(('Text Files','*.txt'),))

        self.__txt_template_file = sg.InputText(default_text = 'C:\Path\To\Word_Doc_Template_file.docx', do_not_clear = True)

        btn_browse_template_file = sg.FileBrowse(file_types=(('Word Files','*.doc*'),))

        btn_apply = sg.Button('Apply')

        self.__txt_updates = sg.Multiline(self.DEFAULT_UPDATE_TEXT, size=(50,15))

        layout = [
            [btn_browse_file_path_doc, self.__txt_file_path_doc],
            [btn_browse_template_file, self.__txt_template_file],
            [btn_apply],
            [self.__txt_updates]
        ]

        self.Layout(layout)

    def start(self, callback):

        while True:

            event, value = self.Read()

            if event is None: break

            elif event == 'Apply':

                if self.__data_valid(): 
                    
                    self.__txt_updates.Update(self.DEFAULT_UPDATE_TEXT)
                    
                    self.__execute_callback(callback)

        self.Close()

    def __data_valid(self):

        self.__file_path_doc = self.__txt_file_path_doc.Get()

        self.__template_file = self.__txt_template_file.Get()

        return (
            (self.__file_path_doc is not None and len(self.__file_path_doc.strip()) > 0) and
            (self.__template_file is not None and len(self.__template_file.strip()) > 0)
        )

    def __execute_callback(self,callback):

        callback(
            self.__file_path_doc, 
            self.__template_file,
            self.update_status_text
            )

        self.__file_path_doc = None

        self.__template_file = None

    def update_status_text(self,text,do_replace=False):

        if do_replace: update_text = text

        else:

            update_text = (
                text
                if self.__txt_updates.Get() == self.DEFAULT_UPDATE_TEXT
                else
                f'{self.__txt_updates.Get()}\n{text}'
            )

        self.__txt_updates.Update(update_text)

        self.Refresh()

