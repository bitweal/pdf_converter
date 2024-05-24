import subprocess
import multiprocessing
import os
import argparse


def worker(func_name, input_file, output_file, extra_args, result_queue):
    command = ["python3", "main.py", func_name, "--input", input_file, "--output", output_file] + extra_args
    try:
        subprocess.run(command, check=True)
        result_queue.put(True)
    except subprocess.CalledProcessError:
        result_queue.put(False)


def run_function_in_parallel(func_name, input_file, extra_args, max_parallelism, output_extension):
    processes = []
    result_queue = multiprocessing.Queue()

    for i in range(max_parallelism):
        output_file = f"media/{func_name}_output_{i}.{output_extension}"
        p = multiprocessing.Process(target=worker, args=(func_name, input_file, output_file, extra_args, result_queue))
        processes.append(p)
        p.start()

    success = True
    for p in processes:
        p.join()
        if not result_queue.get():
            success = False

    return success


def test_function(func_name, input_file, extra_args, max_parallelism, output_extension, max_allowed_parallelism):
    while run_function_in_parallel(func_name, input_file, extra_args, max_parallelism, output_extension):
        max_parallelism += 10
        if max_parallelism > max_allowed_parallelism:
            break
    return max_parallelism - 10


def main():
    parser = argparse.ArgumentParser(description='Test Maximum Parallelism')
    parser.add_argument('--max_parallelism', type=int, default=500, help='Maximum number of parallel processes to test')
    args = parser.parse_args()

    test_pdf = "test.pdf"
    os.makedirs("media", exist_ok=True)

    max_parallelism_results = {}

    # List of functions to test in order with corresponding output extensions
    functions_to_test = [
        ("pdf_to_word", test_pdf, [], "docx"),
        ("word_to_pdf", "media/pdf_to_word_output_0.docx", [], "pdf"),
        ("pdf_to_excel", test_pdf, [], "xlsx"),
        ("excel_to_pdf", "media/pdf_to_excel_output_0.xlsx", [], "pdf"),
        ("pdf_to_jpg", test_pdf, [], "jpg"),
        ("jpg_to_pdf", "output_d", [], "pdf"),
        ("merge", f"{test_pdf},media/excel_to_pdf_output_0.pdf", [], "pdf"),
        ("split", test_pdf, ["--start_page", "0", "--end_page", "1"], "pdf"),
        ("compress", test_pdf, ["--dpi", "100"], "pdf"),
        ("add_page_numbers", test_pdf, ["--position", "middle_bottom"], "pdf"),
        ("protect_pdf", test_pdf, ["--password", "mypassword"], "pdf"),
        ("unlock_pdf", "media/protect_pdf_output_0.pdf", ["--password", "mypassword"], "pdf")
    ]

    for func_name, input_file, extra_args, output_extension in functions_to_test:
        max_parallelism = test_function(func_name, input_file, extra_args, 10, output_extension, args.max_parallelism)
        max_parallelism_results[func_name] = max_parallelism

    for func_name, max_parallelism in max_parallelism_results.items():
        print(f"{func_name} - {max_parallelism} max parallel processes")


if __name__ == "__main__":
    main()
