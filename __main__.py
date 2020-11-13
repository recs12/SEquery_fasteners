# -*- coding: utf-8 -*-

import sys
import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System")
clr.AddReference("System.Runtime.InteropServices")

import System
import System.Runtime.InteropServices as SRI
import System.Collections
from System import Console

from helper import create_query
from query import CriteriaProperties

from SolidEdgeAssembly.QueryConditionConstants import seQueryConditionIs
from SolidEdgeAssembly.QueryConditionConstants import seQueryConditionIsNot
from SolidEdgeAssembly.QueryConditionConstants import seQueryConditionContains
from SolidEdgeAssembly.QueryPropertyConstants import seQueryPropertyReference
from SolidEdgeAssembly.QueryPropertyConstants import seQueryPropertyCustom
from SolidEdgeAssembly.QueryPropertyConstants import seQueryPropertyCategory


# import SolidEdgeAssembly as SEAssembly

__project__ = "query_fasteners"
__author__ = "recs"
__version__ = "0.0.3"
__update__ = "2020-11-13"


def create_various_queries(asm, search_subassemblies):
    # asm = application.ActiveDocument
    print("part: %s\n" % asm.Name)
    assert asm.Type == 3, "This macro only works on .asm document."

    # HARDWARE QUERIES
    # ==================

    # hardware object is created only once because it is used in all the following queries
    hardware = CriteriaProperties(
        seQueryPropertyCategory, "Category", seQueryConditionContains, "HARDWARE"
    )

    # Hardware [PLATED.ZINC]
    zinc = CriteriaProperties(
        seQueryPropertyCustom, "DSC_A", seQueryConditionContains, "ZINC PLATED"
    )
    create_query(
        asm.Queries, "Hardware [PLATED.ZINC]", [
            hardware.criterias, zinc.criterias],
            search_subassemblies
    )

    # Hardware [SS]
    ss = CriteriaProperties(
        seQueryPropertyCustom, "DSC_F", seQueryConditionContains, "SS.3"
    )
    create_query(asm.Queries, "Hardware [SS]", [
                 hardware.criterias, zinc.criterias], 
                 search_subassemblies)

    # Hardware [SS.304]
    ss304 = CriteriaProperties(
        seQueryPropertyCustom, "DSC_F", seQueryConditionContains, "[SS.304]"
    )
    create_query(
        asm.Queries, "Hardware [SS.304]", [hardware.criterias, ss304.criterias], search_subassemblies
    )

    # Hardware [SS.316]
    ss316 = CriteriaProperties(
        seQueryPropertyCustom, "DSC_F", seQueryConditionContains, "[SS.316]"
    )
    create_query(
        asm.Queries, "Hardware [SS.316]", [hardware.criterias, ss316.criterias], search_subassemblies
    )

    # "Hardware INCH"
    inch = CriteriaProperties(
        seQueryPropertyCustom, "JDEPRP1", seQueryConditionIsNot, "Metric Fastener"
    )
    create_query(asm.Queries, "Hardware INCH", [
                 hardware.criterias, inch.criterias],
                 search_subassemblies
                 )

    # "Hardware METRIC"
    metric = CriteriaProperties(
        seQueryPropertyCustom, "JDEPRP1", seQueryConditionContains, "Metric Fastener")
    not_flat_washer = CriteriaProperties(
        seQueryPropertyCustom, "CATEGORY_VB", seQueryConditionIsNot, "FLAT WASHER")
    create_query(asm.Queries, "Hardware METRIC", [
                 hardware.criterias, metric.criterias, not_flat_washer.criterias],
                 search_subassemblies
                 )


def stop():
    sys.exit()


def remove_all_queries(assembly):
    print("queries number: %s" % assembly.Queries.Count)
    # Remove query in the collection of queries
    created_queries = [
        "Hardware [PLATED.ZINC]",
        "Hardware [SS]",
        "Hardware [SS.304]",
        "Hardware [SS.316]",
        "Hardware INCH",
        "Hardware METRIC",
    ]

    for query in created_queries:
        assembly.Queries.Remove(query)
        print("[DELETED] %s " % query)


def would_do_like_to_create_or_remove_all_queries():
    response = raw_input(
    """
    Press [*] to create queries with all parts even those in the subassemblies.
    Press [-] to create queries without the parts in the subassemblies.
    Press [/] to delete all queries.
    """
    ).lower()
    choice = {"*": "create_various_queries_all", "-": "create_various_queries_edited_level", "/": "remove_all_queries"}
    return choice.get(response)


def user_confirmation_to_continue():
    response = raw_input(
        """Would you like to create fasteners queries in the Select Tools? (Press y/[Y] to proceed.)"""
    )
    if response.lower() in ["y", "yes"]:
        pass
    else:
        print("Process canceled")
        sys.exit()


def main():
    try:
        user_confirmation_to_continue()
        answer = would_do_like_to_create_or_remove_all_queries()
        application = SRI.Marshal.GetActiveObject("SolidEdge.Application")
        assembly = application.ActiveDocument

        if answer == "create_various_queries_all":
            create_various_queries(assembly, True)

        elif answer == "create_various_queries_edited_level":
            create_various_queries(assembly, False)

        elif answer == "remove_all_queries":
            remove_all_queries(assembly)

        else:
            pass

    except Exception as ex:
        print(ex)

    finally:
        raw_input("\nPress any key to exit...")
        stop()


if __name__ == "__main__":
    print(
        "%s\n--author:%s --version:%s --last-update :%s\n"
        % (__project__, __author__, __version__, __update__)
    )
    main()
