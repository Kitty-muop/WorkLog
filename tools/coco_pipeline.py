"""
WorkLog CocoIndex Pipeline
Incremental data indexing for WorkLog codebase and documentation.
"""
import os
import sys
from pathlib import Path
import cocoindex

@cocoindex.flow_def(name="WorklogIndex")
def worklog_flow(flow_builder: cocoindex.FlowBuilder, data_scope: cocoindex.DataScope):
    project_root = str(Path(__file__).resolve().parent.parent)
    data_scope["files"] = flow_builder.add_source(
        cocoindex.sources.LocalFile(path=project_root, binary=False)
    )
    
    collector = data_scope.add_collector()
    
    with data_scope["files"].row() as doc:
        collector.collect(
            filename=doc["filename"],
            content=doc["content"]
        )

if __name__ == "__main__":
    print("WorkLog CocoIndex flow definition loaded successfully.")
