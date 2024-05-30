class FlowchartNode:
    def __init__(self, id, visioId, type, title, description, actor=None, decisions=None, to_ids = [], to_visio_ids = []):
        self.id = id
        self.visioId = visioId
        self.type = type
        self.title = title
        self.description = description
        self.actor = actor
        self.decisions = decisions
        self.to_ids = to_ids
        self.to_visio_ids = to_visio_ids

    @staticmethod
    def from_spreadsheet_row(row):
        id = str(row["Step ID"]).strip()
        description = str(row["Description"]).strip()
        actor = str(row["Responsible"]).strip()
        decisions = [s.strip() for s in str(row["Decision"]).strip().split("/")]
        if '' in decisions:
            decisions.remove('')
        if 'nan' in decisions:
            decisions.remove('nan')

        try:
            if len(decisions) > 0:
                type = "Decision"
            else:
                type = "Process"
        except AttributeError:
            pass

        return FlowchartNode(id, None, type, None, description, actor, decisions)

    def is_decision(self):
        return self.type == "Decision"
