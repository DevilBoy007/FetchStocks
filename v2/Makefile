app: ./venv/bin/activate
	python3 app_v2.py
./venv/bin/activate: requirements.txt
	python3 -m venv venv
	pip3 install -r requirements.txt
clean:
	rm -rf venv
	find . -type f -name '*.pyc' -delete
