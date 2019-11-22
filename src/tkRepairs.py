# -*- coding: utf-8 -*-
"""
Created on Wed Nov 20 13:06:13 2019

@author: v.shkaberda
"""

from _version import __version__
from datetime import date, datetime
from decimal import Decimal
from functools import wraps
from tkcalendar import DateEntry
from tkinter import ttk, messagebox
from tkHyperlinkManager import HyperlinkManager
import os
import tkinter as tk


# example of subsription and default recipient
EMAIL_TO = b'\xd0\xa4\xd0\xbe\xd0\xb7\xd0\xb7\xd0\xb8|\xd0\x9b\
\xd0\xbe\xd0\xb3\xd0\xb8\xd1\x81\xd1\x82\xd0\xb8\xd0\xba\xd0\xb0|\xd0\x90\
\xd0\xbd\xd0\xb0\xd0\xbb\xd0\xb8\xd1\x82\xd0\xb8\xd0\xba\xd0\xb8'.decode()


def deco_check_conn(method, *args, **kwargs):
    """ Decorator raises Error when no connection provided.
    """
    @wraps(method)
    def inner(self, *args, **kwargs):
        if not self.conn:
            raise NoConnectionProvidedError(method)
        return method(self, *args, **kwargs)
    return inner


class RepairError(Exception):
    """Base class for exceptions in this module."""
    pass


class NoConnectionProvidedError(RepairError):
    """ Exception raised if no connection provided.

    Attributes:
        expression - input expression in which the error occurred;
        message - explanation of the error.
    """
    def __init__(self, expression, message='Connection is not provided'):
        self.expression = expression
        self.message = message
        super().__init__(self.expression, self.message)


class UpdateRequiredError(Exception):
    """ Exception raised if restart is required.

    Attributes:
        expression - input expression in which the error occurred;
        message - explanation of the error.
    """
    def __init__(self, expression,
                 message='Необходимо обновление приложения'):
        self.expression = expression
        self.message = message
        super().__init__(self.expression, self.message)


class AccessError(tk.Tk):
    """ Raise an error when user doesn't have permission to work with app.
    """
    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showerror(
            'Ошибка доступа',
            'Нет прав для работы с приложением.\n'
            'Для получения прав обратитесь на рассылку ' + EMAIL_TO
        )
        self.destroy()


class LoginError(tk.Tk):
    """ Raise an error when user doesn't have permission to work with db.
    """
    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showerror(
            'Ошибка подключения',
            'Нет прав для работы с сервером.\n'
            'Обратитесь на рассылку ' + EMAIL_TO
        )
        self.destroy()


class NetworkError(tk.Tk):
    """ Raise a message about network error.
    """
    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showerror(
            'Ошибка cети',
            'Возникла общая ошибка сети.\nПовторите попытку позже'
        )
        self.destroy()


class ReinstallRequiredError(tk.Tk):
    """ Raise a message about restart needed after update.
    """
    def __init__(self, *args):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showinfo(
            'Необходимо обновление',
            ('Для возобновления работы необходимо обновить приложение\n' +
             '\n'.join(map(str, args)))
        )
        self.destroy()


class UnexpectedError(tk.Tk):
    """ Raise a message when an unexpected exception occurs.
    """
    def __init__(self, *args):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showinfo(
            'Непредвиденное исключение',
            'Возникло непредвиденное исключение\n' + '\n'.join(map(str, args))
        )
        self.destroy()


class RepairTk(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title('Ремонт техники')
        self.iconbitmap('resources/repairs.ico')
        # handle the window close event
        self.protocol("WM_DELETE_WINDOW", self.quit_with_confirmation)
        # Customize header style (used in PreviewForm)
        style = ttk.Style()
        style.element_create("HeaderStyle.Treeheading.border", "from", "default")
        style.layout("HeaderStyle.Treeview.Heading", [
            ("HeaderStyle.Treeheading.cell", {'sticky': 'nswe'}),
            ("HeaderStyle.Treeheading.border", {'sticky': 'nswe', 'children': [
                ("HeaderStyle.Treeheading.padding", {'sticky': 'nswe', 'children': [
                    ("HeaderStyle.Treeheading.image", {'side': 'right', 'sticky':''}),
                    ("HeaderStyle.Treeheading.text", {'sticky': 'we'})
                ]})
            ]}),
        ])
        style.configure("HeaderStyle.Treeview.Heading",
            background="#eaeaea", foreground="black", relief='groove', font=('Arial', 8))
        style.map("HeaderStyle.Treeview.Heading",
                  relief=[('active', 'sunken'), ('pressed', 'flat')])

        style.map('ButtonGreen.TButton')
        style.configure('ButtonGreen.TButton', foreground='green')

        style.map('ButtonRed.TButton')
        style.configure('ButtonRed.TButton', foreground='red')

        style.configure("TMenubutton", background='white')
        self._center_window(1200, 600)

    def _center_window(self, w, h):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        start_x = int((screen_width/2) - (w/2))
        start_y = int((screen_height/2) - (h*0.55))

        self.geometry('{}x{}+{}+{}'.format(w, h, start_x, start_y))

    def quit_with_confirmation(self):
        if messagebox.askokcancel("Выход", "Выйти из приложения?"):
            self.destroy()


class RepairApp():
    def __init__(self, root, connection, user_info, references):
        """ Initialize app and store all data.

        root: tkinter Tk() instance.

        connection: class that provides connection to db.

        user_info: tuple, contains of:
            - UserID, int;
            - ShortUserName, str;
            - AccessType, int;
            - isSuperUser, bit.

       references: dict of references dicts, {{ref_type1:{ref_name, ref_id},
                                                         {ref_name2, ref_id2},
                                              },
                                              {ref_type2:{ref_name, ref_id}
                                              }
                                             }

        """
        self.root = root
        self.conn = connection
        self.user_info = user_info
        self.UserID = user_info.UserID
        # Filters and other explicitly used info
        self.refs = references
        self._create_refs()
        self.rows = None
        self.sort_reversed_index = None  # reverse sorting for the last sorted column

    def _add_user_label(self, parent):
        """ Adds user name in bottom right corner. """
        user_label = tk.Label(parent, text='Пользователь: ' +
                     self.user_info.ShortUserName + '  Версия ' + __version__,
                     font=('Arial', 8))
        user_label.pack(side=tk.RIGHT, anchor=tk.SE, padx=2)

    def _center_popup_window(self, newlevel, w, h, static_geometry=True):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        start_x = int((screen_width/2) - (w/2))
        start_y = int((screen_height/2) - (h * 0.7))

        if static_geometry is True:
            newlevel.geometry('{}x{}+{}+{}'.format(w, h, start_x, start_y))
        else:
            newlevel.geometry('+{}+{}'.format(start_x, start_y))

    def _clear_filters(self):
        self.createdby_box.set('Все')
        self.status_box.set('Все')

    def _create_refs(self):
        """ Create references used in filters. """
        self.list_mfrs = ('Все', *self.refs['ListMfrs'])
        self.list_rc = ('Все', *self.refs['ListObjects'])
        self.list_owners = ('Все', *self.refs['ListTechnicsOwners'])
        self.list_tech_types = ('Все', *self.refs['ListTechnicsTypes'])
        self.list_people = ('Все', *self.refs['People'])
        self.list_status = ('Все', *self.refs['status_list'])
        self.list_store_types = ('Все', *self.refs['TypeStore'])

    def _format_float(self, sum_float):
        return '{:,.2f}'.format(sum_float).replace(',', ' ').replace('.', ',')

    @deco_check_conn
    def _get_repair_list(self):
        """ Extract information from filters and get repairs list. """
#        filters = {'initiator': self.initiatorsID[self.initiator_box.current()],
#                   'mvz': self.get_mvzSAP(self.mvz_box.get()),
#                   'office': (self.office_box.current() and
#                              self.office[self.office_box.current()]),
#                   'date_type': self.date_type.current(),
#                   'date_m': self.date_entry_m.get_selected(),
#                   'date_y': self.year.get() if self.date_entry_y.get() else 0.,
#                   'sumtotal_from': float(self.sumtotal_from.get_float_form()
#                                          if self.sum_entry_from.get() else 0),
#                   'sumtotal_to': float(self.sumtotal_to.get_float_form()
#                                        if self.sum_entry_to.get() else 0),
#                   'nds':  self.nds.get(),
#                   'statusID': (self.status_box.current() and
#                              self.statusID[self.status_box.current()]),
#                   'payment_num': self.search_by_num.get().strip()
#                   }
        filters = {}
        self.rows = self.conn._get_repair_list(user_info=self.user_info,
                                               **filters)

    def _init_table(self, parent):
        """ Creates treeview. """
        if isinstance(self.headings, dict):
            self.table["columns"] = tuple(self.headings.keys())
            self.table["displaycolumns"] = tuple(k for k, v
                      in self.headings.items() if v > 0)
            for head, width in self.headings.items():
                self.table.heading(head, text=head, anchor=tk.CENTER)
                self.table.column(head, width=6*width, anchor=tk.CENTER,
                                  stretch=False)

        else:
            self.table["columns"] = self.headings
            self.table["displaycolumns"] = self.headings
            for head in self.headings:
                self.table.heading(head, text=head, anchor=tk.CENTER)
                self.table.column(head, width=50*len(head), anchor=tk.CENTER)

        # for debug
        #self._show_rows(rows=((123, 456, 789), ('abc', 'def', 'ghk')))

        msg = 'Incorrect status order for proper table formatting. '
        assert self.list_status[1] == 'Созд.', '{}1!=Созд.'.format(msg)
        assert self.list_status[2] == 'Пров.', '{}2!=Пров.'.format(msg)
        assert self.list_status[3] == 'Удал.', '{}3!=Удал.'.format(msg)
        for tag, bg in zip(self.list_status[1:4],
                           ('#ffffc8', 'lightgreen', '#f66e6e')):
            self.table.tag_configure(tag, background=bg)
        #self.table.tag_configure('oddrow', background='lightgray')

        #self.table.bind('<Double-1>', self._show_detail)
        self.table.bind('<Button-1>', self._sort, True)

        scrolltable_h = tk.Scrollbar(parent, orient='horizontal', command=self.table.xview)
        scrolltable_v = tk.Scrollbar(parent, command=self.table.yview)
        self.table.configure(xscrollcommand=scrolltable_h.set,
                             yscrollcommand=scrolltable_v.set)
        scrolltable_h.pack(side=tk.BOTTOM, fill=tk.X)
        scrolltable_v.pack(side=tk.RIGHT, fill=tk.Y)

    def _make_buttons_frame(self):
        """ Bottom frame with status, version, user info etc.
        """
        buttons_frame = tk.Frame(self.root)

        bt1 = ttk.Button(buttons_frame, text="Создать запись", width=15,
                         command=self._refresh, style='ButtonGreen.TButton')
        bt1.pack(side=tk.LEFT, padx=15, pady=5)

        bt2 = ttk.Button(buttons_frame, text="Создать копию", width=15,
                         command=self.root.destroy)
        bt2.pack(side=tk.LEFT, padx=15, pady=5)

        bt3 = ttk.Button(buttons_frame, text="Удалить запись", width=15,
                         command=self.root.destroy, style='ButtonRed.TButton')
        bt3.pack(side=tk.LEFT, padx=15, pady=5)

        bt4 = ttk.Button(buttons_frame, text="Переместить технику", width=20,
                         command=self.root.destroy)
        bt4.pack(side=tk.LEFT, padx=15, pady=5)

        bt5 = ttk.Button(buttons_frame, text="Добавить технику", width=20,
                         command=self.root.destroy)
        bt5.pack(side=tk.LEFT, padx=15, pady=5)

        bt6 = ttk.Button(buttons_frame, text="Выход", width=15,
                         command=self.root.quit_with_confirmation)
        bt6.pack(side=tk.RIGHT, padx=15, pady=5)

        return buttons_frame

    def _make_main_frame(self):
        """ Main frame.
        """
        main_frame = ttk.LabelFrame(self.root, text=' Список ремонтов ',
                                    name='main_frame')

        # Top Frame with filters
        top_main = tk.Frame(main_frame, name='top_main')
        main_label = tk.Label(top_main, text='Фильтры',
                              padx=10, font=('Arial', 8, 'bold'))

        row1_cf = tk.Frame(top_main, name='row1_cf', padx=15)

        createdby_label = tk.Label(row1_cf, text='Создал', padx=20)
        self.createdby_box = ttk.Combobox(row1_cf, width=20, state='readonly')
        self.createdby_box['values'] = self.list_people

        status_label = tk.Label(row1_cf, text='Статус', padx=20)
        self.status_box = ttk.Combobox(row1_cf, width=10, state='readonly')
        self.status_box['values'] = self.list_status

        # Bottom Frame with table
        bottom_main = tk.Frame(main_frame, name='bottom_main', padx=5)

        # column name and width
        self.headings = {'ID': 8, 'Создал': 20, 'Дата/время создания': 0,
            'StatusID': 0, 'Статус': 10, 'РЦ': 20, 'Склад': 20, 'Наряд-заказ': 14,
            'Собственник': 15, 'Вид техники': 20, 'Производитель': 16,
            'Модель': 12, 'Серийный номер': 20, 'Текущие моточасы': 20,
            'Дата поломки': 23, 'Дата завершения ремонта':25,
            'Описание неисправности': 30, 'Проведённые работы': 30,
            'Кол-во единиц': 14, 'Ед. изм.': 8}
        #self.headings=('a', 'bb', 'cccc')  # for debug

        self.table = ttk.Treeview(bottom_main, show='headings',
                                      selectmode='browse',
                                      style='HeaderStyle.Treeview'
                                      )
        self._init_table(bottom_main)
        self.table.pack(expand=True, fill=tk.BOTH)
        head = self.table["columns"]
        msg = 'Heading order must be reviewed. Wrong heading: '
        assert head[3] == 'StatusID', '{}StatusID'.format(msg)

        main_label.pack(side=tk.TOP, expand=False, anchor=tk.NW)
        createdby_label.pack(side=tk.LEFT)
        self.createdby_box.pack(side=tk.LEFT)
        self.status_box.pack(side=tk.RIGHT)
        status_label.pack(side=tk.RIGHT)
        row1_cf.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        top_main.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        bottom_main.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, padx=5, pady=5)

        return main_frame

    def _make_main_menu(self):
        """ Make top frame with menu.
        """
        main_menu = tk.Frame(self.root)
        file_menu = tk.Menubutton(main_menu, text='Файл',
                                  underline=0, takefocus=0)
        help_menu = tk.Menubutton(main_menu, text='Справка',
                                  underline=0, takefocus=0)
        file_menu.pack(side=tk.LEFT)
        help_menu.pack(side=tk.RIGHT)
        fm = tk.Menu(file_menu, tearoff=0)
        file_menu['menu'] = fm
        hm = tk.Menu(help_menu, tearoff=0)
        help_menu['menu'] = hm

        fm.add_command(label='Выход', underline=0,
                       command=self.root.quit_with_confirmation)
        hm.add_command(label='О программе...', underline=0,
                       command=self._show_about)
        return main_menu

    def _make_status_frame(self):
        """ Bottom frame with status, version, user info etc.
        """
        bottom_frame = tk.Frame(self.root)
        self._add_user_label(bottom_frame)
        return bottom_frame

    def _raise_Toplevel(self, frame, title, width, height,
                        static_geometry=True, options=()):
        """ Create and raise new frame with limits.
        Input:
        frame - class, Frame class to be drawn in Toplevel;
        title - str, window title;
        width - int, width parameter to center window;
        height - int, height parameter to center window;
        static_geometry - bool, if True - width and height will determine size
            of window, otherwise size will be determined dynamically;
        options - tuple, arguments that will be sent to frame.
        """
        newlevel = tk.Toplevel(self.root)
        newlevel.transient(self.root)  # disable minimize/maximize buttons
        newlevel.title(title)
        newlevel.bind('<Escape>', lambda e, w=newlevel: w.destroy())
        frame(newlevel, *options)
        newlevel.resizable(width=False, height=False)
        self._center_popup_window(newlevel, width, height, static_geometry)
        newlevel.focus()
        newlevel.grab_set()

    def _refresh(self):
        """ Refresh repairs information. """
        with self.conn as sql:
            self.rows = sql.get_repair_list()
        self._show_rows(self.rows)

    def _show_about(self, event=None):
        """ Raise frame with info about app. """
        self._raise_Toplevel(frame=AboutFrame,
                             title='Ремонт техники v. ' + __version__,
                             width=400, height=140)

    def _show_rows(self, rows):
        """ Refresh table with new rows. """
        self.table.delete(*self.table.get_children())
        if not rows:
            return
        for row in rows:
            # tag = (Status)
            self.table.insert('', tk.END,
                values=tuple(map(lambda val: self._format_float(val)
                if isinstance(val, Decimal) else '' if val is None else val, row)),
                tags=(row[4],))

    def _sort(self, event):
        if self.table.identify_region(event.x, event.y) == 'heading' and self.rows:
            # determine index of displayed column
            disp_col = int(self.table.identify_column(event.x)[1:]) - 1
            # determine index of this column in self.rows
            sort_col = self.table["columns"].index(self.table["displaycolumns"][disp_col])
            self.rows.sort(key=lambda x: (x[sort_col] is None, x[sort_col]),
                           reverse=self.sort_reversed_index == sort_col)
            # store index of last sorted column if sort wasn't reversed
            self.sort_reversed_index = None if self.sort_reversed_index==sort_col else sort_col
            self._show_rows(self.rows)

    def build(self):
        """ Building app structure.
        """
        frame_menu = self._make_main_menu()
        frame_main = self._make_main_frame()
        frame_buttons = self._make_buttons_frame()
        frame_statusbar = self._make_status_frame()
        frame_menu.pack(side=tk.TOP, fill=tk.X)
        ttk.Separator(self.root).pack(fill=tk.X)
        frame_statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        frame_buttons.pack(side=tk.BOTTOM, fill=tk.X)
        frame_main.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)

    def mainloop(self):
        """ Redirection to root.mainloop.
        """
        self.root.mainloop()

    def run(self):
        """ Function runs app. Build and mainloop could be run manually,
        for example, if you need to add another function to root.after:
            app.build()
            root.after(200, some_func)
            app.mainloop()
        """
        try:
            self.build()
        except Exception as e:
            self.root.destroy()
            raise
        self._clear_filters()
        self.root.after(200, self._refresh)
        self.mainloop()


class AboutFrame(tk.Frame):
    """ Creates a frame with copyright and info about app. """
    def __init__(self, parent):
        super().__init__(parent)

        self.top = ttk.LabelFrame(self, name='top_af')

        logo = tk.PhotoImage(file='resources/repairs.png')
        self.logo_label = tk.Label(self.top, image=logo)
        self.logo_label.image = logo  # keep a reference to avoid gc!

        self.copyright_text = tk.Text(self.top, bg='#f1f1f1',
                                      font=('Arial', 8), relief=tk.FLAT)
        hyperlink = HyperlinkManager(self.copyright_text)

        self.copyright_text.insert(tk.INSERT,
                                  'Ремонт техники v. ' + __version__ +'\n')
        self.copyright_text.insert(tk.INSERT, "\n")

        def link_license():
            path = 'resources\\LICENSE.txt'
            os.startfile(path)

        self.copyright_text.insert(tk.INSERT,
                            'Copyright © 2019 Офис контроллинга логистики\n')
        self.copyright_text.insert(tk.INSERT, 'MIT License',
                                   hyperlink.add(link_license))

        self.bt = ttk.Button(self, text="Закрыть", width=10,
                        command=parent.destroy)
        self.pack_all()

    def pack_all(self):
        self.bt.pack(side=tk.BOTTOM, pady=2)
        self.top.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10)
        self.logo_label.pack(side=tk.LEFT, padx=10, pady=2)
        self.copyright_text.pack(side=tk.LEFT, padx=10, pady=10)
        self.pack(fill=tk.BOTH, expand=True, pady=5)


if __name__ == '__main__':
    from db_connect import DBConnect
    from collections import namedtuple

    UserInfo = namedtuple('UserInfo', ['UserID', 'ShortUserName',
                                       'AccessType', 'isSuperUser'])
    conn = DBConnect(server='s-kv-center-s59', db='AnalyticReports')
    root = RepairTk()
    app = RepairApp(root=root,
                    connection=conn,
                    user_info=UserInfo(24, 'TestName', 1, 1),
                    references = {}
                    )
    app.run()
