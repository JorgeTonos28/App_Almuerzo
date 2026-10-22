import os

def read_file(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        return f.read()

def main():
    try:
        index = read_file('index.html')
        css = read_file('css.html')
        js = read_file('js.html')

        assembled = index.replace("<?!= include('css'); ?>", css)
        assembled = assembled.replace("<?!= include('js'); ?>", js)

        with open('mock_app.html', 'w', encoding='utf-8') as f:
            f.write(assembled)
        print("Created mock_app.html")
    except Exception as e:
        print(f"Error assembling app: {e}")

if __name__ == '__main__':
    main()
