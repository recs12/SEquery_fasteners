class CriteriaProperties(object):
    """ Role of reducing the number of variable to pass in the function create_query()
    """

    def __init__(self, query_property, property, query_condition, condition):
        self.query_property = query_property
        self.property = property
        self.query_condition = query_condition
        self.condition = condition

    @property
    def criterias(self):
        return (
            self.query_property,
            self.property,
            self.query_condition,
            self.condition,
        )
