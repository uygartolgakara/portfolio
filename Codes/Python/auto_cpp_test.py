# -*- coding: utf-8 -*-
"""
Produced on Fri Jul 21 15:50:11 2023
@author: KUY3IB
"""
import subprocess
import unittest


class TestCPlusPlusCode(unittest.TestCase):

    def test_cpp_program(self):
        # The C++ program to test
        program = "./test_program"

        # If your program takes input, you can pass it in here
        input_data = "5"

        # Run the program and get the output
        process = subprocess.run(
            [program], input=input_data, text=True, capture_output=True)

        # Check the return code (0 usually means success)
        self.assertEqual(process.returncode, 0)

        # Check the output of the program
        output = process.stdout.strip()
        self.assertEqual(output, "10")


if __name__ == "__main__":
    unittest.main()
