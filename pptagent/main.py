from utils import Planner, Parser, Processor, Applier, Reporter, SharedLogMemory
def main(self, user_input):
    self.planner = Planner()
    self.parser = Parser()
    self.processor = Processor()
    self.applier = Applier()
    self.reporter = Reporter()

    self.memory = SharedLogMemory()

    plan = self.planner(user_input)
    to_process = self.parser(plan)
    to_apply = self.processor(to_process)
    result = self.applier(to_apply)

    to_report = self.memory(plan, to_process, to_apply, result)
    
    self.reporter(to_report)

