using Dapper;
using Microsoft.Extensions.Options;
using MySqlConnector;

namespace WordStyleCheckService.Worker;

public class Db(IOptionsMonitor<Options> options)
{
    MySqlConnection GetConnection()
    {
        return new MySqlConnection(options.CurrentValue.ConnectionString);
    }

    public class DequeuedTaskInfo
    {
        public uint PendingTaskId { get; set; }
        public uint TaskId { get; set; }
        public string TaskType { get; set; }
        public string TaskData { get; set; }
    }
    
    public class TaskInfo
    {
        public string TaskType { get; set; }
        public TaskCompletionStatus CompletionStatus { get; set; }
        public string? ResultData { get; set; }
    }

    public async Task<TaskInfo?> GetTask(uint taskId)
    {
        using var conn = GetConnection();

        return await conn.QueryFirstOrDefaultAsync<TaskInfo>("""
            SELECT TaskType, CompletionStatus, ResultData
            FROM bl_task
            WHERE TaskId = @taskId
        """, new {taskId});
    }

    public async Task<uint> EnqueueTask(string taskType, string taskData)
    {
        using var conn = GetConnection();
        uint taskId = await conn.QueryFirstAsync<uint>("""
            INSERT INTO bl_task
            (TaskType, TaskData)
            VALUES 
            (@taskType, @taskData)
            RETURNING TaskId
            """,
            new { taskType, taskData });

        await conn.ExecuteAsync("""
            INSERT INTO bl_pendingtask
            (TaskId, ExpireTs)
            VALUES
            (@taskId, TIMESTAMPADD(DAY, 1, NOW()))
        """, new { taskId });
        
        return taskId;
    }
    
    public async Task<DequeuedTaskInfo?> TryDequeuePendingTask()
    {
        using var conn = GetConnection();
        return await conn.QueryFirstOrDefaultAsync<DequeuedTaskInfo>("""
            START TRANSACTION;
            SET @pendId = -1;
            SELECT `PendingTaskId` INTO @pendId FROM `bl_pendingtask` WHERE ProcessingStatus = 1 ORDER BY PendingTaskId ASC LIMIT 1 FOR UPDATE;
            UPDATE bl_pendingtask SET `ProcessingStatus` = 2, `ExpireTs` = TIMESTAMPADD(SECOND, 30, NOW()) WHERE `PendingTaskId` = @pendId;
            SELECT `PendingTaskId`, `TaskId`, `TaskType`, `TaskData` 
            	FROM `bl_pendingtask` pt 
            	JOIN  `bl_task` USING (`TaskId`)
                WHERE `PendingTaskId` = @pendId;
            COMMIT;
            """);
    }

    public async Task ReportTaskSuccess(uint pendingTaskId, string result)
    {
        using var conn = GetConnection();
        await conn.ExecuteAsync("""
            START TRANSACTION;
            UPDATE `bl_task` t JOIN `bl_pendingtask` pt USING (`TaskId`)
                SET t.`CompletionStatus` = 2, t.`ResultData` = @result
                WHERE pt.`PendingTaskId` = @pendingTaskId;
            DELETE FROM `bl_pendingtask` WHERE `PendingTaskId` = @pendingTaskId;
            COMMIT;
            """, new { pendingTaskId, result });
    }

    public async Task ReportTaskFailure(uint pendingTaskId, string error)
    {
        using var conn = GetConnection();
        await conn.ExecuteAsync("""
            START TRANSACTION;
            UPDATE `bl_task` t JOIN `bl_pendingtask` pt USING (`TaskId`)
                SET t.`CompletionStatus` = 3, t.`ResultData` = @error
                WHERE pt.`PendingTaskId` = @pendingTaskId;
            DELETE FROM `bl_pendingtask` WHERE `PendingTaskId` = @pendingTaskId;
            COMMIT;
            """, new { pendingTaskId, error });
    }


    public async Task MaintainTaskQueue(string errorExpiredNotTaken, string errorProcessingTimeout)
    {
        using var conn = GetConnection();
        await conn.ExecuteAsync("""
            START TRANSACTION;
            SET SQL_SAFE_UPDATES=0;
            UPDATE `bl_task` t 
                JOIN `bl_pendingtask` pt USING (`TaskId`) 
                SET 
            		t.`ResultData` = "errorExpiredNotTaken", 
            		t.`CompletionStatus` = 3,
                    pt.`ProcessingStatus` = 4
                WHERE pt.`ProcessingStatus` = 1 AND pt.ExpireTs < NOW();
            DELETE FROM `bl_pendingtask` WHERE `ProcessingStatus` = 4;

            UPDATE `bl_task` t 
                JOIN `bl_pendingtask` pt USING (`TaskId`) 
                SET 
            		t.`ResultData` = "errorProcessingTimeout", 
            		t.`CompletionStatus` = 3,
                    pt.`ProcessingStatus` = 5
                WHERE pt.`ProcessingStatus` = 2 AND pt.ExpireTs < NOW();
            DELETE FROM `bl_pendingtask` WHERE `ProcessingStatus` = 5;
            SET SQL_SAFE_UPDATES=1;
            COMMIT;
            """);
    }


    public enum PendingTaskProcessingStatus 
    {
        Null = 0,
        Created = 1,
        Taken = 2,
        Success = 3,
        ExpiredNotTaken = 4,
        ProcessingTimeout = 5
    }


    public enum TaskCompletionStatus
    {
        Null = 0,
        Created = 1,
        Success = 2,
        Failed = 3 
    }

}