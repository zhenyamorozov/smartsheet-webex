
import smartsheet

ss_api = smartsheet.Smartsheet()

ss_api.errors_as_exceptions(True)

ss_sheet = ss_api.Sheets.get_sheet("qxHP4HJfpGRc9mgxJCm29VXw7vGF4M2vcgPcx5g1")

columns = ss_sheet.columns

for col in columns:
    print()
    print(col)
    pass

pass



