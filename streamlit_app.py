import runpy, sys, os

app_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'AI', '연간 발주 계획 새우기')
sys.path.insert(0, app_dir)
os.chdir(app_dir)
runpy.run_path(os.path.join(app_dir, 'web_app.py'), run_name='__main__')
