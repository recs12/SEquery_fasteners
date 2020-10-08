# -*- coding: utf-8 -*-
__VERSION__ = "0.0.0"

import sys
import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System")
clr.AddReference("System.Runtime.InteropServices")

import System
import System.Collections
import System.Runtime.InteropServices as SRI
from System import Console

# import SolidEdgeAssembly as SEAssembly

from SolidEdgeAssembly.QueryPropertyConstants import seQueryPropertyCategory
from SolidEdgeAssembly.QueryPropertyConstants import seQueryPropertyCustom
from SolidEdgeAssembly.QueryPropertyConstants import seQueryPropertyReference
from SolidEdgeAssembly.QueryConditionConstants import seQueryConditionContains
from SolidEdgeAssembly.QueryConditionConstants import seQueryConditionIsNot
from SolidEdgeAssembly.QueryConditionConstants import seQueryConditionIs

from query import CriteriaProperties
from helper import create_query


def application():
    return SRI.Marshal.GetActiveObject("SolidEdge.Application")


def active_document(application):
    return application.ActiveDocument


def create_various_queries(asm):
    # asm = application.ActiveDocument
    print("part: %s\n" % asm.Name)
    assert asm.Type == 3, "This macro only works on .asm"

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
        asm.Queries, "Hardware [PLATED.ZINC]", [hardware.criterias, zinc.criterias]
    )

    # Hardware [SS]
    ss = CriteriaProperties(
        seQueryPropertyCustom, "DSC_F", seQueryConditionContains, "SS.3"
    )
    create_query(asm.Queries, "Hardware [SS]", [hardware.criterias, zinc.criterias])

    # Hardware [SS.304]
    ss304 = CriteriaProperties(
        seQueryPropertyCustom, "DSC_F", seQueryConditionContains, "[SS.304]"
    )
    create_query(
        asm.Queries, "Hardware [SS.304]", [hardware.criterias, ss304.criterias]
    )

    # Hardware [SS.316]
    ss316 = CriteriaProperties(
        seQueryPropertyCustom, "DSC_F", seQueryConditionContains, "[SS.316]"
    )
    create_query(
        asm.Queries, "Hardware [SS.316]", [hardware.criterias, ss316.criterias]
    )

    # "Hardware INCH"
    inch = CriteriaProperties(
        seQueryPropertyCustom, "JDEPRP1", seQueryConditionIsNot, "Metric Fastener"
    )
    create_query(asm.Queries, "Hardware INCH", [hardware.criterias, inch.criterias])

    # "Hardware METRIC"
    metric = CriteriaProperties(seQueryPropertyCustom, "JDEPRP1", seQueryConditionContains, "Metric Fastener")
    not_flat_washer = CriteriaProperties(seQueryPropertyCustom, "CATEGORY_VB", seQueryConditionIsNot, "FLAT WASHER")
    create_query(asm.Queries, "Hardware METRIC", [hardware.criterias, metric.criterias, not_flat_washer.criterias])


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


def would_do_like_to_create_or_removeall_queries():
    response = raw_input(
        """Create queries[Q] or delete[D] all queries? (Press [Q] or [D] to proceed...)"""
    ).lower()
    choice = {"q": create_various_queries, "d": remove_all_queries}
    return choice.get(response)


def user_confirmation_to_continue():
    response = raw_input(
        """Create fasteners/reference queries? (Press y/[Y] to proceed.)"""
    )
    if response.lower() in ["y", "yes"]:
        pass
    else:
        print("Process canceled")
        sys.exit()


def main():
    try:
        print(__VERSION__)
        user_confirmation_to_continue()
        answer = would_do_like_to_create_or_removeall_queries()
        app = application()
        assembly = active_document(app)
        if answer is create_various_queries:
            create_various_queries(assembly)
        elif answer is remove_all_queries:
            remove_all_queries(assembly)
        else:
            pass

    except Exception as ex:
        print(ex)

    finally:
        raw_input("\nPress any key to exit...")
        stop()



if __name__ == "__main__":
    main()
