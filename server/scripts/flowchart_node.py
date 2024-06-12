class FlowchartNode:
    def __init__(self, id, visioId, type, title, description, actor=None, decisions=None, jump_to_ids=None, nest_level=1):
        self.id = id
        self.visioId = visioId
        self.type = type
        self.title = title
        self.description = description
        self.actor = actor
        self.decisions = decisions
        self.jump_to_ids = jump_to_ids  # Ids for jump connections
        self.nest_level = nest_level

    def add_jump(self, to_id):
        self.jump_to_ids.append(to_id)

    @staticmethod
    def from_spreadsheet_row(row):
        print("ROW IS")
        print(row)
        id = str(row["Step ID"]).strip()
        print("ID IS", id)
        description = str(row["Description"]).strip()
        print("DESCRIPTION IS", description)
        actor = str(row["Responsible"]).strip()
        print("ACTOR IS", actor)
        decisions = [s.strip()
                     for s in str(row["Decision"]).strip().split("/")]
        print("DECISIONS ARE", decisions)
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

        fcn = FlowchartNode(id, None, type, None, description, actor, decisions, [])
        print("FCN IS", fcn)
        print("FCN DECISIONS ARE", fcn.decisions)
        print("FCN DESCRIPTION IS", fcn.description)
        return fcn

    def is_decision(self):
        return self.type == "Decision"
