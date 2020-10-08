from SolidEdgeAssembly.QueryScopeConstants import seQueryScopeAllParts


def create_query(queries, query_name, criterias, search_subassemblies=True):
    """Create a query in select tools panel.
    """

    # Add the query here:
    query = queries.Add(query_name)
    query.Scope = seQueryScopeAllParts
    query.SearchSubassemblies = search_subassemblies

    # loop throught the criterias
    for criteria in criterias:
        query.AddCriteria(criteria[0], criteria[1], criteria[2], criteria[3])
    print(
        "[QUERY]: Created: {0:.<25}{1:.>50}".format(
            query_name, query.MatchesCount.ToString()
        )
    )
