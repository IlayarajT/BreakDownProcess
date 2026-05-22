class ProgressBar:
    def print(self, iteration, total, prefix='', suffix='', decimals=1, length=100, fill='█', printEnd="\r"):
        percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
        filledLength = int(length * iteration // total)
        bar = fill * filledLength + ' ' * (length - filledLength)
        print(f'\r{prefix} {bar} {suffix}', end=printEnd)
        if iteration == total:
            print()
