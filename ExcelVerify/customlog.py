import time

class CustomLog:

    def __init__(self):
        self.content = []
        self.create_time = time.time()
        self.unknown_error = "\nUnknown Error!\n"
        self.status_text = {
            "file":{
                "import": "\nReading <{0}>...\n",
                "fail": "{0} FAILED!\n",
                "success": "{0} done.\n",
                "not_exist": "\"{0}\" doesn't exist!\n",
                "all_exist": "{0} all exist!\n",
                "not_contain": "\"{0}\" doesn't contain available data!\n",
            },

            "target":{
                "import": "\nReading <{0}>...\n",
                "not_exist": "{0} does not exist! \n",
                "duplicate": "WARNING!!! Duplicated string detected in {0}! \n",
                "no_duplicate": "No duplicate in {0}! \n",
                "success": "Getting {0} successfully!\n",
                "matched": "All matched for {0}.\n",
                "not_found": "\"{0}\" is not found. \n",
                "is_duplicate": "{0} is duplicate!\n",
            },

            "other":{
                "section_start": "\n---- {0} BEGINS ----",
                "section_finished": "---- {0} ENDS ----\n",
                "error_count": "{0} id(s) are not matched. \n",
                "empty": "Keywords of {0} should not be empty!\n",
                "invalid_check_type": "The check type of {0} is invalid!\n",
                "bundle_miss": "{0} is missing!!!\n",
            }
        }

    def log(self, log):
        self.content.append(log)
        print(log)

    def write_log(self, filename):
        try:
            f = open(filename, "w+")
        except IOError:
            print("Can not write to log file \"{0}\"! \n".format(filename))
        else:
            log_time_string = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(self.create_time))
            f.write("Start checking on {0}".format(log_time_string))

            for log in self.content:
                f.write(log)
            f.close()
            print("Logs are written into {0} \n".format(filename))

    def log_status(self, log_type, target_name, status):
        # Types: file, target, other
        try:
            log = self.status_text[log_type][status]
        except:
            log = self.unknown_error
        else:
            self.log(log.format(target_name))
