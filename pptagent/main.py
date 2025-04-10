from utils import Planner, Parser, Processor, Applier, Reporter, SharedLogMemory
def main(self, user_input):
    self.planner = Planner()
    self.parser = Parser()
    self.processor = Processor()
    self.applier = Applier()
    self.reporter = Reporter()

    self.memory = SharedLogMemory()

    # 계획 세우기
    plan = self.planner(user_input)
    
    # ppt에서 데이터 파싱하기
    to_process = self.parser(plan)
    
    # 데이터 processing
    to_apply = self.processor(to_process)
    
    # ppt에 적용하고, True return하기
    result:bool = self.applier(to_apply)

    # 이전 내용 모두 넣어서 간추리고 사용자에게 report할 내용 return
    to_report = self.memory(plan, to_process, to_apply, result)
    
    # 사용자에게 report하기기
    self.reporter(to_report)

