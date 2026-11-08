"""Account for every visible paragraph change by its exact authorized text span."""
from paperfit import q


class TextIntegrity:
    def __init__(self, stories):
        self.stories = stories
        self.expected = {p: ''.join(p.itertext(q('t'))) for root in stories.values() for p in root.iter(q('p'))}
        self.edits = 0

    def replace(self, paragraph, start, before, after):
        text = self.expected[paragraph]
        if text[start:start+len(before)] != before:
            raise ValueError('文本变更与已核对的原位置不一致')
        self.expected[paragraph] = text[:start] + after + text[start+len(before):]
        self.edits += before != after

    def new(self, paragraph, text):
        self.expected[paragraph] = text

    def verify(self):
        actual = {p: ''.join(p.itertext(q('t'))) for root in self.stories.values() for p in root.iter(q('p'))}
        if actual != self.expected:
            raise ValueError('检测到授权编号或引用以外的文字或段落变更，停止输出')
        return {'paragraphs_checked':len(actual),'authorized_text_edits':self.edits}
