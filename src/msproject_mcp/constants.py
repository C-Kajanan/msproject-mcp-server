"""Microsoft Project COM enumeration values used by the server.

Values come from the Project VBA object model reference. They are hard-coded
so the server does not depend on a generated COM type library (makepy).
"""

# PjTaskLinkType
LINK_TYPES = {
    "FF": 0,  # pjFinishToFinish
    "FS": 1,  # pjFinishToStart
    "SF": 2,  # pjStartToFinish
    "SS": 3,  # pjStartToStart
}
LINK_TYPE_NAMES = {v: k for k, v in LINK_TYPES.items()}

# PjResourceTypes
RESOURCE_TYPES = {
    "work": 0,  # pjResourceTypeWork
    "material": 1,  # pjResourceTypeMaterial
    "cost": 2,  # pjResourceTypeCost
}
RESOURCE_TYPE_NAMES = {v: k for k, v in RESOURCE_TYPES.items()}

# PjConstraint
CONSTRAINT_TYPES = {
    "as_soon_as_possible": 0,  # pjASAP
    "as_late_as_possible": 1,  # pjALAP
    "must_start_on": 2,  # pjMSO
    "must_finish_on": 3,  # pjMFO
    "start_no_earlier_than": 4,  # pjSNET
    "start_no_later_than": 5,  # pjSNLT
    "finish_no_earlier_than": 6,  # pjFNET
    "finish_no_later_than": 7,  # pjFNLT
}
CONSTRAINT_TYPE_NAMES = {v: k for k, v in CONSTRAINT_TYPES.items()}

# PjSaveType (FileClose)
PJ_DO_NOT_SAVE = 0
PJ_SAVE = 1

# PjDocExportType (DocumentExport)
PJ_PDF = 0

# FileSaveAs format strings
SAVE_FORMATS = {
    "mpp": "MSProject.MPP",
    "mpt": "MSProject.MPT",
    "xml": "MSProject.XML",
}
