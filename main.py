import logging
import logging.handlers
import os

from conversion_940 import *
from conversion_944 import *


path = "C:\\FTP\\GPAEDIProduction\\TG1-SAG\\In\\"
files = os.listdir(path)


def main():
    for file in files:
        filename = str(file).split("_")
        rem_extension = file.split(".")
        try:
            try:
                if rem_extension[1] == "csv":
                    conversion = Convert_940(path + file)
                    conversion.parse_csv()
                    os.replace(path + file, path + "\\Archive\\940\\" + file)
                elif rem_extension[1] == "xml":
                    conversion = Convert_944(path + file)
                    conversion.parseXMLemail()
                    os.replace(path + file, "C:\\FTP\\GPAEDIProduction\\TG1-SAG\\In\\Archive\\944\\" +
                               rem_extension[0] + '_' + datetime.now().strftime("%Y%m%d%H%M%S") + ".xml")
            except BaseException:
                logger = logging.getLogger()
                fileHandler = logging.FileHandler(
                    "C:\\FTP\\GPAEDIProduction\\TG1-SAG\\Logs\\" + filename[0] + "_" + filename[1] + "_" + datetime.now().strftime("%Y%m%d") + ".log")
                smtp_handler = logging.handlers.SMTPHandler(mailhost=("smtp.office365.com", 587),
                                                            fromaddr="noreply@gpalogisticsgroup.com",
                                                            toaddrs=["cmaya@gpalogisticsgroup.com", "avelazquez@gpalogisticsgroup.com", "gpaops20@gpalogisticsgroup.com"],
                                                            subject=filename[0] + " failed to process for client " + filename[1],
                                                            credentials=('noreply@gpalogisticsgroup.com', 'Turn*17300'),
                                                            secure=())
                formatter = logging.Formatter('%(asctime)s - %(name)s - %(message)s')
                fileHandler.setFormatter(formatter)
                logger.addHandler(fileHandler)
                logger.addHandler(smtp_handler)
                logger.exception("An exception was triggered")
                os.replace(path + file, path + "\\err_" + file)
        except IndexError:
            pass
        except PermissionError:
            pass


if __name__ == '__main__':
    main()

