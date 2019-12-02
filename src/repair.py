# -*- coding: utf-8 -*-
"""
Created on Thu Nov 21 17:28:04 2019

@author: v.shkaberda
"""
from _version import version_info
from collections import defaultdict, namedtuple
from db_connect import DBConnect
from log_error import writelog
from pyodbc import Error as SQLError
import sys
import tkRepairs as tkr


def main():
    # Reading config and setting parameters
    config = {}
    try:
        with open('config.ini', 'r') as f:
            for line in f:
                # Ignore lines started with "#" and blank lines
                if not line.startswith('#') and line.strip():
                    k = line[:line.index(':')].strip()
                    v = line[line.index(':')+1:].strip()
                    config[k] = v

    except ValueError:
        writelog('Error: config.ini have unappropriate lines: ' + line)
        sys.exit(1)

    conn = DBConnect(server=config['server'],
                     db=config['db'])
    UserInfo = namedtuple('UserInfo', ['UserID', 'ShortUserName',
                                       'AccessType', 'isSuperUser'])
    refs = defaultdict(dict)
    try:
        with conn as sql:
            if sql.get_version_for_comparison() != version_info[:2]:
                raise tkr.UpdateRequiredError(version_info[:2])

            access_permitted = sql.access_check()
            if not access_permitted:
                tkr.AccessError()
                sys.exit()

            # load references
            user_info = UserInfo(*sql.get_user_info())
            for ref_type, ref_id, ref_name in sql.get_references():
                refs[ref_type][ref_name] = ref_id

        # Run app
        root = tkr.RepairTk()
        app = tkr.RepairApp(root=root,
                            connection=conn,
                            user_info=user_info,
                            references=refs
                            )
        app.run()

    except SQLError as e:
        # login failed
        if e.args[0] in ('28000', '42000'):
            tkr.LoginError()
        else:
            raise

    except tkr.UpdateRequiredError as e:
        writelog(e)
        tkr.ReinstallRequiredError()


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        writelog(e)
    finally:
        sys.exit()
