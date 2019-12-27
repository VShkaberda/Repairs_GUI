# -*- coding: utf-8 -*-
"""
Created on Wed Nov 20 13:06:13 2019

@author: v.shkaberda
"""

from _version import __version__
from autocomplete_entry import AutocompleteEntry
from datetime import datetime
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

class StringSumVar(tk.StringVar):
    """ Contains function that returns var formatted in a such way, that
        it can be converted into a float without an error.
    """
    def get_float_form(self, *args, **kwargs):
        return (super().get(*args, **kwargs).replace(' ', '')
                                            .replace('\n', '')
                                            .replace(',', '.'))


class RepairTk(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title('Ремонт техники')
        self.iconbitmap('resources/repairs.ico')
        # handle the window close event
        self.protocol("WM_DELETE_WINDOW", self.quit_with_confirmation)
        # extend Ctrl+C/V/X to other keyboards
        self.bind_all("<Key>", self._onKeyRelease, '+')
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

    def _onKeyRelease(*args):
        event = args[1]
        # check if Ctrl pressed
        if not (event.state == 12 or event.state == 14):
            return
        if event.keycode == 88 and event.keysym.lower() != 'x':
            event.widget.event_generate("<<Cut>>")
        elif event.keycode == 86 and event.keysym.lower() != 'v':
            event.widget.event_generate("<<Paste>>")
        elif event.keycode == 67 and event.keysym.lower() != 'c':
            event.widget.event_generate("<<Copy>>")

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
        self.rc_box.set('Все')
        self.store_box.set('Все')
        self.status_box.set('Все')
        self.owner_box.set('Все')
        self.mfr_box.set('Все')
        self.tech_type_box.set('Все')

    def _create_refs(self):
        """ Create references used in filters. """
        self.list_mfrs = ('Все', *self.refs.get('ListMfrs', ()))
        self.list_rc = ('Все', *self.refs.get('ListObjects', ()))
        self.list_owners = ('Все', *self.refs.get('ListTechnicsOwners', ()))
        self.list_tech_types = ('Все', *self.refs.get('ListTechnicsTypes', ()))
        self.list_people = ('Все', *self.refs.get('People', ()))
        self.list_status = ('Все', *self.refs['status_list'])
        self.list_store_types = ('Все', *self.refs.get('TypeStore', ()))

    def _format_float(self, sum_float):
        return '{:,.2f}'.format(sum_float).replace(',', ' ').replace('.', ',')

    @deco_check_conn
    def _get_repair_list(self):
        """ Extract information from filters and get repairs list. """
        filters = {
            'created_by': self.refs['People'].get(self.createdby_box.get(), None),
            'rc': None if self.rc_box.get() == 'Все' else self.rc_box.get(),
            'store': self.refs['TypeStore'].get(self.store_box.get(), None),
            'status': self.refs['status_list'].get(self.status_box.get(), None),
            'owner': self.refs['ListTechnicsOwners'].get(self.owner_box.get(), None),
            'mfr': self.refs['ListMfrs'].get(self.mfr_box.get(), None),
            'tech_type': self.refs['ListTechnicsTypes'].get(self.tech_type_box.get(), None)
        }
        with self.conn as sql:
            self.rows = sql.get_repair_list(**filters)

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
        assert self.list_status[2] == 'Фикс.', '{}2!=Фикс.'.format(msg)
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

    @deco_check_conn
    def _load_refs(self):
        """ Load references for creating-editing repairs. """
        with self.conn as sql:
            tech_info = sql.get_technics_info()
            tech_info = dict((data[0], data[1:]) for data in tech_info)
            measure_units = sql.get_measure_units()
            measure_units = dict((data[0], data[1]) for data in measure_units)
            objects = sql.get_objects()
            objects = dict((data[1:], data[0]) for data in objects)
        options = {'conn': self.conn,
                   'userID': self.UserID,
                   'tech_info': tech_info,
                   'measure_units': measure_units,
                   'objects': objects}
        return options

    def _make_buttons_frame(self):
        """ Bottom frame with status, version, user info etc.
        """
        buttons_frame = tk.Frame(self.root)

        bt1 = ttk.Button(buttons_frame, text="Создать запись", width=15,
                         command=self._popup_create_form,
                         style='ButtonGreen.TButton')
        bt1.pack(side=tk.LEFT, padx=15, pady=5)

        bt2 = ttk.Button(buttons_frame, text="Создать копию", width=15,
                         command=self._popup_create_copy_form)
        bt2.pack(side=tk.LEFT, padx=15, pady=5)

        bt3 = ttk.Button(buttons_frame, text="Редактировать", width=15,
                         command=self.root.destroy)
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

        row1_cf = tk.Frame(top_main, name='row1_cf')

        createdby_label = tk.Label(row1_cf, text='Создал', width=12, anchor=tk.E)
        self.createdby_box = ttk.Combobox(row1_cf, width=20, state='readonly')
        self.createdby_box['values'] = self.list_people

        rc_label = tk.Label(row1_cf, text='РЦ', width=16, anchor=tk.E)
        self.rc_box = ttk.Combobox(row1_cf, width=20, state='readonly')
        self.rc_box['values'] = self.list_rc

        store_label = tk.Label(row1_cf, text='Склад', width=16, anchor=tk.E)
        self.store_box = ttk.Combobox(row1_cf, width=20, state='readonly')
        self.store_box['values'] = self.list_store_types

        status_label = tk.Label(row1_cf, text='Статус', padx=40)
        self.status_box = ttk.Combobox(row1_cf, width=10, state='readonly')
        self.status_box['values'] = self.list_status

        row2_cf = tk.Frame(top_main, name='row2_cf')

        owner_label = tk.Label(row2_cf, text='Владелец', width=12, anchor=tk.E)
        self.owner_box = ttk.Combobox(row2_cf, width=20, state='readonly')
        self.owner_box['values'] = self.list_owners

        mfr_label = tk.Label(row2_cf, text='Производитель', width=16, anchor=tk.E)
        self.mfr_box = ttk.Combobox(row2_cf, width=20, state='readonly')
        self.mfr_box['values'] = self.list_mfrs

        tech_type_label = tk.Label(row2_cf, text='Тип техники', width=16, anchor=tk.E)
        self.tech_type_box = ttk.Combobox(row2_cf, width=20, state='readonly')
        self.tech_type_box['values'] = self.list_tech_types

        bt02 = ttk.Button(row2_cf, text="Очистить", width=15,
                          command=self._clear_filters)
        bt02.pack(side=tk.RIGHT)

        bt01 = ttk.Button(row2_cf, text="Обновить", width=15,
                          command=self._refresh)
        bt01.pack(side=tk.RIGHT, padx=20)

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
        assert head[0] == 'ID', '{}ID'.format(msg)
        assert head[3] == 'StatusID', '{}StatusID'.format(msg)

        main_label.pack(side=tk.TOP, expand=False, anchor=tk.NW)
        createdby_label.pack(side=tk.LEFT, padx=5)
        self.createdby_box.pack(side=tk.LEFT, padx=10)
        rc_label.pack(side=tk.LEFT)
        self.rc_box.pack(side=tk.LEFT, padx=10)
        store_label.pack(side=tk.LEFT)
        self.store_box.pack(side=tk.LEFT, padx=10)
        self.status_box.pack(side=tk.RIGHT)
        status_label.pack(side=tk.RIGHT)
        row1_cf.pack(side=tk.TOP, fill=tk.X, padx=20, pady=10)
        owner_label.pack(side=tk.LEFT, padx=5)
        self.owner_box.pack(side=tk.LEFT, padx=10)
        mfr_label.pack(side=tk.LEFT)
        self.mfr_box.pack(side=tk.LEFT, padx=10)
        tech_type_label.pack(side=tk.LEFT)
        self.tech_type_box.pack(side=tk.LEFT, padx=10)
        row2_cf.pack(side=tk.TOP, fill=tk.X, padx=20, pady=5)
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
                       command=self._popup_about)
        return main_menu

    def _make_status_frame(self):
        """ Bottom frame with status, version, user info etc.
        """
        bottom_frame = tk.Frame(self.root)
        self._add_user_label(bottom_frame)
        return bottom_frame

    def _popup_about(self, event=None):
        """ Raise frame with info about app. """
        self._raise_Toplevel(frame=AboutFrame,
                             title='Ремонт техники v. ' + __version__,
                             width=400, height=140)

    def _popup_create_copy_form(self, event=None):
        curRow = self.table.focus()
        if not curRow:
            return
        options = self._load_refs()
        options['current_repairID'] = self.table.item(curRow).get('values')[0]
        self._raise_Toplevel(frame=CreateCopyFrame,
                             title='Данные о ремонте',
                             width=800, height=400,
                             static_geometry=False,
                             refresh_after=True,
                             options=options)

    def _popup_create_form(self, event=None):
        options = self._load_refs()
        self._raise_Toplevel(frame=CreateFrame,
                             title='Данные о ремонте',
                             width=800, height=400,
                             static_geometry=False,
                             refresh_after=True,
                             options=options)

    def _raise_Toplevel(self, frame, title, width, height,
                        static_geometry=True, refresh_after=False, options={}):
        """ Create and raise Toplevel with Frame.
        Input:
        frame - class, Frame class to be drawn in Toplevel;
        title - str, window title;
        width - int, width parameter to center window;
        height - int, height parameter to center window;
        static_geometry - bool, if True - width and height will determine size
            of window, otherwise size will be determined dynamically;
        refresh_after - bool, if True - call self._refresh() after Toplevel
            is destroyed;
        options - dictionary, kwargs that will be sent to frame.
        """
        newlevel = tk.Toplevel(self.root)
        newlevel.title(title)
        newlevel.bind('<Escape>', lambda e, w=newlevel: w.destroy())
        frame(newlevel, **options)
        newlevel.resizable(width=False, height=False)
        self._center_popup_window(newlevel, width, height, static_geometry)
        newlevel.focus()
        newlevel.grab_set()
        if refresh_after:
            self.root.wait_window(newlevel)
            self._refresh()

    def _refresh(self):
        """ Refresh repairs information. """
        self._get_repair_list()
        self._show_rows(self.rows)

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


class CreateFrame(tk.Frame):
    def __init__(self, parent, conn, userID, tech_info, measure_units, objects):
        # hide until all frames have been created
        parent.withdraw()
        super().__init__(parent)
        self.parent = parent
        self.conn = conn
        self.UserID = userID
        self.tech_info = tech_info
        self.measure_units = measure_units
        self.measure_units_list = (*self.measure_units,)
        self.objects = objects
        self.allowed_objects = None
        # parameter to control SN correctness to prevent _create execution
        self.is_sn_correct = False

        frame_name = self._get_frame_name()
        main_label = tk.Label(self, text=frame_name, width=20,
                              anchor=tk.CENTER, font=('Arial', 12, 'bold'))
        main_label.grid(row=0, column=0, columnspan=4, pady=10, padx=20)

        sn_label = tk.Label(self, text='Серийный номер',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        sn_label.grid(row=1, column=0, pady=5, padx=20, sticky=tk.E)

        self.sn_entry = AutocompleteEntry(list(self.tech_info.keys()),
                                     self, width=30)
        self.sn_entry.bind("<FocusOut>", self._check_SN)
        self.sn_entry.grid(row=1, column=1, pady=5, padx=5)

        outfitorder_label = tk.Label(self, text='Наряд-заказ',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        outfitorder_label.grid(row=1, column=2, pady=5, padx=20, sticky=tk.EW)

        self.outfitorder = tk.StringVar()
        outfitorder_entry = tk.Entry(self, width=30,
                                     textvariable=self.outfitorder)
        outfitorder_entry.grid(row=1, column=3, pady=5, padx=5)

        tech_type_label = tk.Label(self, text='Вид техники',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        tech_type_label.grid(row=2, column=0, pady=5, padx=20, sticky=tk.E)

        self.tech_type = tk.StringVar()
        tech_type_entry = tk.Entry(self, width=30, state='readonly', takefocus=0,
                                   textvariable=self.tech_type)
        tech_type_entry.grid(row=2, column=1, pady=5, padx=5)

        model_label = tk.Label(self, text='Модель',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        model_label.grid(row=2, column=2, pady=5, padx=20, sticky=tk.EW)

        self.model = tk.StringVar()
        model_entry = tk.Entry(self, width=30, state='readonly', takefocus=0,
                               textvariable=self.model)
        model_entry.grid(row=2, column=3, pady=5, padx=5)

        owner_label = tk.Label(self, text='Владелец',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        owner_label.grid(row=3, column=0, pady=5, padx=20, sticky=tk.E)

        self.owner = tk.StringVar()
        owner_entry = tk.Entry(self, width=30, state='readonly', takefocus=0,
                               textvariable=self.owner)
        owner_entry.grid(row=3, column=1, pady=5, padx=5)

        mfr_label = tk.Label(self, text='Производитель',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        mfr_label.grid(row=3, column=2, pady=5, padx=20, sticky=tk.EW)

        self.mfr = tk.StringVar()
        mfr_entry = tk.Entry(self, width=30, state='readonly', takefocus=0,
                             textvariable=self.mfr)
        mfr_entry.grid(row=3, column=3, pady=5, padx=5)

        date_broken_label = tk.Label(self, text='Дата поломки',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        date_broken_label.grid(row=5, column=0, pady=5, padx=20, sticky=tk.E)

        self.date_broken = tk.StringVar()
        self.date_broken_entry = DateEntry(self, width=23, state='disabled',
            background='#2888e8', foreground='white', selectbackground ='#2888e8',
            font=('Arial', 9), selectmode='day', borderwidth=2, locale='ru_RU',
            textvariable=self.date_broken)
        self.date_broken.set('')
        self.date_broken_entry.bind("<FocusOut>", self._get_rc_store_owner)
        #self.date_broken.trace("w", self._get_rc_store_owner)
        self.date_broken_entry.grid(row=5, column=1, pady=5, padx=5)

        workhours_label = tk.Label(self, text='Текущие моточасы',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        workhours_label.grid(row=5, column=2, pady=5, padx=20, sticky=tk.E)

        self.workhours = StringSumVar()
        workhours_entry = tk.Entry(self, width=30,
                                   textvariable=self.workhours)
        workhours_entry.grid(row=5, column=3, pady=5, padx=5)

        rc_label = tk.Label(self, text='РЦ',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        rc_label.grid(row=6, column=0, pady=5, padx=20, sticky=tk.E)

        self.rc = tk.StringVar()
        self.rc_box = ttk.Combobox(self, width=27, state='disabled',
                                   textvariable=self.rc)
        self.rc_box.grid(row=6, column=1, pady=5, padx=5)

        store_label = tk.Label(self, text='Склад',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        store_label.grid(row=6, column=2, pady=5, padx=20, sticky=tk.EW)

        self.store = tk.StringVar()
        self.store_box = ttk.Combobox(self, width=27, state='disabled',
                                   textvariable=self.store)
        self.store_box.grid(row=6, column=3, pady=5, padx=5)

        number_units_label = tk.Label(self, text='Кол-во запчастей',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        number_units_label.grid(row=7, column=0, pady=5, padx=20, sticky=tk.E)

        self.number_units = StringSumVar()
        number_units_entry = tk.Entry(self, width=30,
                                      textvariable=self.number_units)
        number_units_entry.grid(row=7, column=1, pady=5, padx=5)

        units_measure_label = tk.Label(self, text='Единица измерения',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        units_measure_label.grid(row=7, column=2, pady=5, padx=20, sticky=tk.EW)

        self.units_measure = tk.StringVar()
        self.units_measure_box = ttk.Combobox(self, width=27, state='readonly',
                                              textvariable=self.units_measure)
        self.units_measure_box['values'] = self.measure_units_list
        self.units_measure.set('не указано')
        self.units_measure_box.grid(row=7, column=3, pady=5, padx=5)

        faultdesc_label = tk.Label(self, text='Описание неисправности',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        faultdesc_label.grid(row=8, column=0, pady=5, padx=20, sticky=tk.E)

        self.faultdesc = tk.StringVar()
        faultdesc_entry = tk.Entry(self, width=89,
                                   textvariable=self.faultdesc)
        faultdesc_entry.grid(row=8, column=1, columnspan=3, pady=5)

        perfwork_label = tk.Label(self, text='Проведённые работы',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        perfwork_label.grid(row=9, column=0, pady=5, padx=20, sticky=tk.E)

        self.perfwork = tk.StringVar()
        perfwork_entry = tk.Entry(self, width=89,
                                  textvariable=self.perfwork)
        perfwork_entry.grid(row=9, column=1, columnspan=3, pady=5)

        date_repair_end_label = tk.Label(self, text='Дата окончания ремонта',
                              anchor=tk.E, font=('Arial', 8, 'bold'))
        date_repair_end_label.grid(row=10, column=1, pady=5, padx=20, sticky=tk.EW)

        self.date_repair_end_active = tk.IntVar()
        date_repair_end_cbox = tk.Checkbutton(self, anchor=tk.W,
                                              variable=self.date_repair_end_active,
                                              command=self._toggle_date_repair_end)
        date_repair_end_cbox.grid(row=10, column=2, pady=5, sticky=tk.W)
        self.date_repair_end_active.set(1)

        self.date_repair_end = tk.StringVar()
        self.date_repair_end_entry = DateEntry(self, width=9, state='readonly',
            background='#2888e8', foreground='white', selectbackground ='#2888e8',
            font=('Arial', 9), selectmode='day', borderwidth=2, locale='ru_RU',
            textvariable=self.date_repair_end)
        self.date_repair_end_entry.grid(row=10, column=2, pady=5, padx=10)

        self._make_buttons()

        self.pack(fill=tk.BOTH, expand=True, padx=30, pady=5)
        # restore after withdraw
        parent.deiconify()

    def _check_SN(self, event):
        self.date_broken.set('')
        self.owner.set('')
        self.rc.set('')
        self.store.set('')
        sn = self.sn_entry.get()
        objs = (self.tech_type, self.model, self.mfr)
        # determine widget name focused next
        new_focus = str(self.focus_get()).split('!')[-1]
        # ignore check for inner listbox / clear and exit buttons
        # include 'None' to ignore check if user clicked outside CreateFrame
        if not sn or new_focus in ('None', 'lbox', 'button3', 'button4'):
            self.date_broken_entry.configure(state='disabled')
            for obj in objs:
                obj.set('')
            return
        try:
            for obj, val in zip(objs, self.tech_info[sn]):
                obj.set(val)
            self.date_broken_entry.configure(state='readonly')
            self.is_sn_correct = True
        except KeyError:
            messagebox.showinfo(
                'Серийный номер',
                'Серийный номер не найден'
            )
            for obj in objs:
                obj.set('')
            self.sn_entry.focus_set()
            self.sn_entry.event_generate('<<SelectAll>>')
            self.date_broken_entry.configure(state='disabled')
            self.is_sn_correct = False

    def _clear_all(self):
        self.sn_entry.var.set('')
        self.outfitorder.set('')
        self.tech_type.set('')
        self.model.set('')
        self.owner.set('')
        self.mfr.set('')
        self.date_broken.set('')
        self.date_broken_entry.configure(state='disabled')
        self.workhours.set('')
        self.rc.set('')
        self.store.set('')
        self.number_units.set('')
        self.units_measure.set('не указано')
        self.faultdesc.set('')
        self.perfwork.set('')
        self.date_repair_end.set('')
        self.is_sn_correct = False

    # deprecated
    def _clear_calendar(self, event, wname):
        if wname == 'date_broken':
            self.date_broken.set('')
        elif wname == 'date_repair_end':
            self.date_repair_end.set('')

    def _convert_date(self, date_in, output=None):
        """ Take date_in and convert it into output format.
            If output is None datetime object is returned.

            date: str in format '%d[./]%m[./]%y' or '%d[./]%m[./]%Y'.
            output: str or None, output format.
        """
        date_in = date_in.replace('/', '.')
        try:
            dat = datetime.strptime(date_in, '%d.%m.%y')
        except ValueError:
            dat = datetime.strptime(date_in, '%d.%m.%Y')
        if output:
            return dat.strftime(output)
        return dat

    def _create(self, is_fixed=False):
        """ Create repair according to filled fields.
            is_fixed - bool, determines StatusID of repair.
        """
        # since sn is checked in _check_SN using <FocusOut> of self.sn_entry
        # just prevent execution istead of raising message
        if not self.is_sn_correct:
            return
        messagetitle = 'Данные о ремонте'
        is_validated = self._validate_repair_info(messagetitle)
        if not is_validated:
            return
        repair_info = self._get_repair_info()
        statusID = 2 if is_fixed else 1
        if is_fixed:
            confirmed = messagebox.askyesno(title='Подтвердите фиксацию',
                message='Создать заявку без возможности редактировать в будущем?')
            if not confirmed:
                return
        with self.conn as sql:
            created_success = sql.create_repair(userID=self.UserID,
                                                **repair_info,
                                                statusID=statusID)
        if created_success == 1:
            messagebox.showinfo(
                    messagetitle, 'Данные внесены'
            )
            self.parent.destroy()
        else:
            messagebox.showerror(
                    messagetitle, 'Произошла ошибка'
            )

    def _get_frame_name(self):
        return 'Данные о ремонте'

    def _get_repair_info(self):
        """ Takes info from filled fields and returns it as dictionary
             in format appropriated for server.
        """
        workhours = self.workhours.get_float_form()
        number_units = self.number_units.get_float_form()
        units_m = self.units_measure.get()
        date_broken = self.date_broken_entry.get_date()
        date_broken = datetime.combine(date_broken, datetime.min.time())
        repair_info = {
            'SN': self.sn_entry.get().strip(),
            'date_broken': self._convert_date(self.date_broken.get()),
            'date_repair_finished': (self._convert_date(self.date_broken.get())
                if self.date_repair_end_active.get() else None),
            'OutfitOrder': self.outfitorder.get().strip() or None,
            'WorkingHours': (float(workhours) if workhours else None),
            'UnitOfMeasureID': self.measure_units[units_m] if units_m else None,
            'NumberOfUnits': (float(number_units) if number_units else None),
            'FaultDescription': self.faultdesc.get().strip() or None,
            'PerformedWork': self.perfwork.get().strip() or None
            }
        return repair_info

    def _get_rc_store_owner(self, *args, **kwargs):
        """ Shows rc, store and owner corresponding to chosen date and SN.
        """
        sn = self.sn_entry.get()
        new_focus = str(self.focus_get()).split('!')[-1]
        # do nothing if user clicked on calendar or clear/exit buttons
        # include 'None' to ignore check if user clicked outside CreateFrame
        if new_focus in ('None', 'calendar', 'button3', 'button4'):
            self.owner.set('')
            self.rc.set('')
            self.store.set('')
            return
        try:
            date_broken = self._convert_date(self.date_broken.get())
        except Exception as e:
            return
        if not date_broken:
            return
        with self.conn as sql:
            self.allowed_objects = sql.get_object_owner_info(sn, date_broken)
        if len(self.allowed_objects) == 0:
            messagebox.showinfo(title='Нет привязки',
                message=('В указанную дату техника не привязана ни к одному РЦ')
                                )
        else:
            self.owner.set(self.allowed_objects[0][0])
            self.rc.set(self.allowed_objects[0][1])
            self.store.set(self.allowed_objects[0][2])

    def _make_buttons(self):
        bt1 = ttk.Button(self, text='Сохранить', width=26,
                         command=self._create, style='ButtonGreen.TButton')
        bt1.grid(row=11, column=0, pady=10)

        bt2 = ttk.Button(self, text='Сохранить (зафиксировать)', width=26,
                         command=lambda: self._create(is_fixed=True),
                         style='ButtonGreen.TButton')
        bt2.grid(row=11, column=1, pady=10)

        bt3 = ttk.Button(self, text='Очистить всё', width=15,
                         command=self._clear_all)
        bt3.grid(row=11, column=2, pady=10)

        self.bt4 = ttk.Button(self, text='Закрыть', width=15,
                         command=self.parent.destroy)
        self.bt4.grid(row=11, column=3, pady=10, padx=10, sticky=tk.E)

    def _toggle_date_repair_end(self):
        """ Toggle states of self.date_repair_end_entry.
        """
        if self.date_repair_end_active.get():
            self.date_repair_end_entry.configure(state="readonly")
        else:
            self.date_repair_end_entry.configure(state="disabled")

    def _validate_repair_info(self, messagetitle):
        """ Validate correctness of filled fields. Returns bool.
        """
        try:
            float(self.workhours.get_float_form() or 0)
        except ValueError:
            messagebox.showerror(
                    messagetitle, 'Некорректно заполнено: Моточасы'
            )
            return False
        try:
            number_units = float(self.number_units.get_float_form() or 0)
        except ValueError:
            messagebox.showerror(
                    messagetitle, 'Некорректно заполнено: Кол-во запчастей'
            )
            return False
        if number_units and not self.units_measure.get():
            messagebox.showerror(
                    messagetitle, 'Не указана ед. измерения запчастей'
            )
            return False
        if not self.rc.get():
            messagebox.showerror(
                    messagetitle, 'Не указано РЦ'
            )
            return False
        if not self.store.get():
            messagebox.showerror(
                    messagetitle, 'Не указан склад'
            )
            return False
        date_validation = self._validate_date_broken()
        if date_validation == 'incorrect_date':
            messagebox.showerror(
                    messagetitle,
                    'Дата поломки некорректна или не заполнена'
            )
            return False
        elif date_validation != 'correct_date':
            messagebox.showerror(title='Ошибка',
                            message=('Возникло непредвиденное исключение\n{}'
                                     .format(date_validation))
                                )
            return False
        return True

    def _validate_date_broken(self):
        """ Validate date correctness. """
        try:
            date_broken = self.date_broken.get()
            self._convert_date(date_broken)
            return 'correct_date'
        except ValueError:
            return 'incorrect_date'
        except Exception as e:
            return e


class CreateCopyFrame(CreateFrame):
    def __init__(self, parent, conn, userID, tech_info, measure_units, objects,
                 current_repairID):
        super().__init__(parent, conn, userID, tech_info, measure_units,
                         objects)
        self.current_repairID = current_repairID
        with self.conn as sql:
            current_repair_info = sql.get_current_repair(self.current_repairID)
        self._set_current_repair(*current_repair_info[0])

    def _set_current_repair(self, sn_entry, outfitorder, tech_type, model,
                            owner, mfr, date_broken, workhours, rc, store,
                            number_units, units_measure, faultdesc, perfwork,
                            date_repair_end):
        self.sn_entry.var.set(sn_entry)
        # next 3 lines imitate behaviour of _check_SN
        self.sn_entry.destroy_listbox()
        self.date_broken_entry.configure(state='readonly')
        self.is_sn_correct = True
        self.outfitorder.set(outfitorder)
        self.tech_type.set(tech_type)
        self.model.set(model)
        self.owner.set(owner)
        self.mfr.set(mfr)
        self.date_broken_entry.set_date(date_broken)
        self.workhours.set(workhours)
        self.rc.set(rc)
        self.store.set(store)
        self.number_units.set(number_units)
        self.units_measure.set(units_measure)
        self.faultdesc.set(faultdesc)
        self.perfwork.set(perfwork)
        if date_repair_end:
            self.date_repair_end_entry.set_date(date_repair_end)
        else:
            self.date_repair_end_active.set(0)
            self.date_repair_end_entry.configure(state="disabled")


class UpdateRepairFrame(CreateCopyFrame):
    def __init__(self, parent, conn, userID, tech_info, measure_units, objects,
                 repairID):
        super().__init__(parent, conn, userID, tech_info, measure_units,
                         objects)
        self.repairID = repairID

    def _make_buttons(self):
        bt1 = ttk.Button(self, text='Сохранить', width=26,
                         command=self._create, style='ButtonGreen.TButton')
        bt1.grid(row=11, column=0, pady=10)

        bt2 = ttk.Button(self, text='Сохранить (зафиксировать)', width=26,
                         command=lambda: self._create(is_fixed=True),
                         style='ButtonGreen.TButton')
        bt2.grid(row=11, column=1, pady=10)

        bt3 = ttk.Button(self, text='Очистить всё', width=15,
                         command=self._clear_all, state='disabled')
        bt3.grid(row=11, column=2, pady=10)

        bt4 = ttk.Button(self, text='Закрыть', width=15,
                         command=self.parent.destroy)
        bt4.grid(row=11, column=3, pady=10, padx=10, sticky=tk.E)


if __name__ == '__main__':
    from collections import namedtuple

    UserInfo = namedtuple('UserInfo', ['UserID', 'ShortUserName',
                                       'AccessType', 'isSuperUser'])
    root = RepairTk()
    app = RepairApp(root=root,
                    connection=None,
                    user_info=UserInfo(24, 'TestName', 1, 1),
                    references = {'status_list': {'Созд.': 1,
                                                  'Фикс.': 2,
                                                  'Удал.': 3}}
                    )
    app.run()
