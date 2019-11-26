# -*- coding: utf-8 -*-
"""
Created on Wed Nov 20 11:40:43 2019

@author: v.shkaberda
"""
from functools import wraps
from tkRepairs import NetworkError
import pyodbc


def monitor_network_state(method):
    """ Show error message in case of network error.
    """
    @wraps(method)
    def wrapper(self, *args, **kwargs):
        try:
            return method(self, *args, **kwargs)
        except pyodbc.Error as e:
            # Network error
            if e.args[0] in ('01000', '08S01', '08001'):
                NetworkError()
    return wrapper


class DBConnect(object):
    """ Provides connection to database and functions to work with server.
    """
    def __init__(self, *, server, db):
        self._server = server
        self._db = db
        # Connection properties
        self.conn_str = (
            'Driver={{SQL Server}};'
            'Server={0};'
            'Database={1};'
            'Trusted_Connection=yes;'.format(self._server, self._db)
        )

    def __enter__(self):
        self.__db = pyodbc.connect(self.conn_str)
        self.__cursor = self.__db.cursor()
        return self

    def __exit__(self, type, value, traceback):
        self.__db.close()

    @monitor_network_state
    def add_movement(self, UserID, ID, SN, ObjectID, date_movement):
        """ Executes procedure that adds info about technics movement.
        """
        query = '''
        exec technics.ADD_MOVEMENT  @UserID = ?,
                                    @ID = ?,
                                    @SN = ?,
                                    @ObjectID = ?,
                                    @date_movement = ?
            '''
        try:
            self.__cursor.execute(query, UserID, ID, SN, ObjectID,
                                  date_movement)
            request_success = self.__cursor.fetchone()[0]
            self.__db.commit()
            return request_success
        except pyodbc.ProgrammingError:
            return

    @monitor_network_state
    def create_repair(self, userID, SN, ObjectID, date_broken,
                      date_repair_finished, OutfitOrder, WorkingHours,
                      UnitOfMeasureID, NumberOfUnits, FaultDescription,
                      PerformedWork, StatusID):
        """ Executes procedure that creates new repair.
        """
        query = '''
        exec technics.CREATE_REPAIR @UserID = ?,
                                    @SN  = ?,
                                    @ObjectID = ?,
                                    @date_broken = ?,
                                    @date_repair_finished = ?,
                                    @OutfitOrder = ?,
                                    @WorkingHours = ?,
                                    @UnitOfMeasureID = ?,
                                    @NumberOfUnits = ?,
                                    @FaultDescription = ?,
                                    @PerformedWork = ?,
                                    @StatusID = ?
            '''
        try:
            self.__cursor.execute(query, userID, SN, ObjectID, date_broken,
                      date_repair_finished, OutfitOrder, WorkingHours,
                      UnitOfMeasureID, NumberOfUnits, FaultDescription,
                      PerformedWork, StatusID)
            request_success = self.__cursor.fetchone()[0]
            self.__db.commit()
            return request_success
        except pyodbc.ProgrammingError:
            return

    @monitor_network_state
    def get_repair_list(self, *, created_by, rc, store, owner, mfr,
                        tech_type, status):
        """ Executes procedure and return repair list.
        """
        query = '''
        exec technics.get_repair_list @created_by = ?,
                                      @rc = ?,
                                      @store = ?,
                                      @owner = ?,
                                      @mfr = ?,
                                      @tech_type = ?,
                                      @status = ?

        '''
        self.__cursor.execute(query, created_by, rc, store, owner, mfr,
                              tech_type, status)
        return self.__cursor.fetchall()

    @monitor_network_state
    def access_check(self):
        """ Check user permission.
            If access permitted returns True, otherwise None.
        """
        self.__cursor.execute("exec [technics].[Access_Check]")
        access = self.__cursor.fetchone()
        # check AccessType and isSuperUser
        if access and (access[0] in (1, 2, 3) or access[1]):
            return True

    @monitor_network_state
    def get_objects(self):
        """ Returns all references used in filters.
        """
        self.__cursor.execute("exec [technics].[get_objects]")
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_references(self):
        """ Returns all references used in filters.
        """
        self.__cursor.execute("exec [technics].[get_references]")
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_user_info(self):
        """ Returns information about current user based on ORIGINAL_LOGIN().
        """
        self.__cursor.execute("exec [technics].[get_user_info]")
        return self.__cursor.fetchone()

    @monitor_network_state
    def get_version_for_comparison(self):
        """ Return version supposed to be the same as client version.
        """
        self.__cursor.execute("exec [technics].[get_version_for_comparison]")
        return tuple(self.__cursor.fetchone())

    @monitor_network_state
    def raw_query(self, query):
        """ Takes the query and returns output from db.
        """
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def update_repair(self, UserID, RepairID, SN, ObjectID, date_broken,
                      date_repair_finished, OutfitOrder, WorkingHours,
                      UnitOfMeasureID, NumberOfUnits, FaultDescription,
                      PerformedWork, StatusID):
        """ Executes procedure that updates repair.
        """
        query = '''
        exec technics.UPDATE_REPAIR @UserID = ?,
                                    @RepairID = ?,
                                    @SN  = ?,
                                    @ObjectID = ?,
                                    @date_broken = ?,
                                    @date_repair_finished = ?,
                                    @OutfitOrder = ?,
                                    @WorkingHours = ?,
                                    @UnitOfMeasureID = ?,
                                    @NumberOfUnits = ?,
                                    @FaultDescription = ?,
                                    @PerformedWork = ?,
                                    @StatusID = ?
            '''
        try:
            self.__cursor.execute(query, UserID, RepairID, SN, ObjectID, date_broken,
                      date_repair_finished, OutfitOrder, WorkingHours,
                      UnitOfMeasureID, NumberOfUnits, FaultDescription,
                      PerformedWork, StatusID)
            request_success = self.__cursor.fetchone()[0]
            self.__db.commit()
            return request_success
        except pyodbc.ProgrammingError:
            return


if __name__ == '__main__':
    with DBConnect(server='s-kv-center-s64', db='CB') as sql:
        query = 'select 42'
        assert sql.raw_query(query)[0][0] == 42, 'Server returns no output.'
    print('Connected successfully.')
    input('Press Enter to exit...')
