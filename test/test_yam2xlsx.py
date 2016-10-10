import os
from openpyxl import load_workbook
from yaml2xlsx.yaml2xlsx import YAML2XLSX


class test_yaml2xlsx(object):
    
    @classmethod
    def setup_class(cls):
        yamltest =\
'''
---
additional_column: [1,2,3,4]
contact:
  - Peter, Tester <p.tester@a.admain.org>
  - Potor Tostor <o.totstor@bdomain.org>
anticipated FTEs: 0.4
time-horizon: 1 year
previous SXS engagement: Ralf Rollypop
project: testproj_1
notes: |
  Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus
  et ante ut ipsum malesuada finibus. Pellentesque quis pretium est,
  et tincidunt magna. Nulla eget purus egestas, gravida nisi egestas,
  imperdiet felis. Maecenas laoreet congue consectetur. Praesent
  pellentesque accumsan risus, a pulvinar diam. Quisque at congue
  quam. Mauris ullamcorper tempus ante, cursus varius massa
  condimentum suscipit. Integer a mi lacus. Phasellus egestas sem ac
  justo rhoncus gravida. Curabitur tellus felis, viverra in purus a,
  bibendum tincidunt erat. Mauris posuere dolor at lorem ultrices, ac
  eleifend tortor eleifend.
---
time-horizon: 8
project: testproj_2
contact:
  - Peter_1, Tester_1 <p.tester_1@a.admain.org>
  - Potor_1 Tostor_1 <o.totstor_1@bdomain.org>
anticipated FTEs: 0.9
another_column: HAHAHA
previous SXS engagement: Mike Muttersohn
notes: |
  Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus
  et ante ut ipsum malesuada finibus. Pellentesque quis pretium est,
  et tincidunt magna. Nulla eget purus egestas, gravida nisi egestas,
  imperdiet felis. Maecenas laoreet congue consectetur. Praesent
  pellentesque accumsan risus, a pulvinar diam. Quisque at congue
  quam. Mauris ullamcorper tempus ante, cursus varius massa
  condimentum suscipit. Integer a mi lacus. Phasellus egestas sem ac
  justo rhoncus gravida. Curabitur tellus felis, viverra in purus a,
  bibendum tincidunt erat. Mauris posuere dolor at lorem ultrices, ac
  eleifend tortor eleifend.
'''
        cls.xlsxlist = [
            (u'notes',
             u'project',
             u'contact',
             u'additional_column',
             u'previous SXS engagement',
             u'time-horizon',
             u'anticipated FTEs',
             u'another_column'),
            (u'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus\net ante ut ipsum malesuada finibus. Pellentesque quis pretium est,\net tincidunt magna. Nulla eget purus egestas, gravida nisi egestas,\nimperdiet felis. Maecenas laoreet congue consectetur. Praesent\npellentesque accumsan risus, a pulvinar diam. Quisque at congue\nquam. Mauris ullamcorper tempus ante, cursus varius massa\ncondimentum suscipit. Integer a mi lacus. Phasellus egestas sem ac\njusto rhoncus gravida. Curabitur tellus felis, viverra in purus a,\nbibendum tincidunt erat. Mauris posuere dolor at lorem ultrices, ac\neleifend tortor eleifend.\n',
            u'testproj_1',
            u"['Peter, Tester <p.tester@a.admain.org>', 'Potor Tostor <o.totstor@bdomain.org>']",
             u'[1, 2, 3, 4]',
             u'Ralf Rollypop',
             u'1 year',
             u'0.4',
             None),
            (u'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus\net ante ut ipsum malesuada finibus. Pellentesque quis pretium est,\net tincidunt magna. Nulla eget purus egestas, gravida nisi egestas,\nimperdiet felis. Maecenas laoreet congue consectetur. Praesent\npellentesque accumsan risus, a pulvinar diam. Quisque at congue\nquam. Mauris ullamcorper tempus ante, cursus varius massa\ncondimentum suscipit. Integer a mi lacus. Phasellus egestas sem ac\njusto rhoncus gravida. Curabitur tellus felis, viverra in purus a,\nbibendum tincidunt erat. Mauris posuere dolor at lorem ultrices, ac\neleifend tortor eleifend.\n',
             u'testproj_2',
             u"['Peter_1, Tester_1 <p.tester_1@a.admain.org>', 'Potor_1 Tostor_1 <o.totstor_1@bdomain.org>']",
             None,
             u'Mike Muttersohn',
             u'8',
             u'0.9',
             u'HAHAHA')]

        with open('testyaml.yml', 'w') as f:
            f.writelines(yamltest)
        cls.C = YAML2XLSX(['testyaml.yml', 'testxlsx.xlsx'])
    
    @classmethod
    def teardown_class(cls):
        os.remove('testyaml.yml')
        os.remove('testxlsx.xlsx')
        
    def test_argparse(self):
        assert(self.C.parameters == {'infile': 'testyaml.yml',
                                     'outfile': 'testxlsx.xlsx'})

    def test_collect_columns(self):
        data = self.C.read_file()
        columns = self.C.collect_columns()
        assert(sorted(columns) == sorted([
            'notes',
            'project',
            'contact',
            'additional_column',
            'previous SXS engagement',
            'time-horizon',
            'anticipated FTEs',
            'another_column']))
        
    def test_yaml2xlsx(self):
        self.C.yaml2xlsx()
        res = list(load_workbook('testxlsx.xlsx').active.values)
        assert(res == self.xlsxlist)
