-- 开启事务，确保数据迁移原子性（要么全成功，要么全回滚）
BEGIN TRANSACTION;

BEGIN TRY
    -- 插入数据：显式映射所有字段，跳过自增列（Id）
    INSERT INTO [RecordSystemV26].[dbo].[SourceData] (
        [ReportedTime],    -- 对应源表createdtime
        [LastChange],      -- 用源表创建时间填充，也可保留默认值sysdatetime()
        [Type],            -- 对应源表type
        [PH],              -- 对应源表PH
        -- 以下是cell1-cell150字段（按顺序映射）
        [Cell1],[Cell2],[Cell3],[Cell4],[Cell5],[Cell6],[Cell7],[Cell8],[Cell9],[Cell10],
        [Cell11],[Cell12],[Cell13],[Cell14],[Cell15],[Cell16],[Cell17],[Cell18],[Cell19],[Cell20],
        [Cell21],[Cell22],[Cell23],[Cell24],[Cell25],[Cell26],[Cell27],[Cell28],[Cell29],[Cell30],
        [Cell31],[Cell32],[Cell33],[Cell34],[Cell35],[Cell36],[Cell37],[Cell38],[Cell39],[Cell40],
        [Cell41],[Cell42],[Cell43],[Cell44],[Cell45],[Cell46],[Cell47],[Cell48],[Cell49],[Cell50],
        [Cell51],[Cell52],[Cell53],[Cell54],[Cell55],[Cell56],[Cell57],[Cell58],[Cell59],[Cell60],
        [Cell61],[Cell62],[Cell63],[Cell64],[Cell65],[Cell66],[Cell67],[Cell68],[Cell69],[Cell70],
        [Cell71],[Cell72],[Cell73],[Cell74],[Cell75],[Cell76],[Cell77],[Cell78],[Cell79],[Cell80],
        [Cell81],[Cell82],[Cell83],[Cell84],[Cell85],[Cell86],[Cell87],[Cell88],[Cell89],[Cell90],
        [Cell91],[Cell92],[Cell93],[Cell94],[Cell95],[Cell96],[Cell97],[Cell98],[Cell99],[Cell100],
        [Cell101],[Cell102],[Cell103],[Cell104],[Cell105],[Cell106],[Cell107],[Cell108],[Cell109],[Cell110],
        [Cell111],[Cell112],[Cell113],[Cell114],[Cell115],[Cell116],[Cell117],[Cell118],[Cell119],[Cell120],
        [Cell121],[Cell122],[Cell123],[Cell124],[Cell125],[Cell126],[Cell127],[Cell128],[Cell129],[Cell130],
        [Cell131],[Cell132],[Cell133],[Cell134],[Cell135],[Cell136],[Cell137],[Cell138],[Cell139],[Cell140],
        [Cell141],[Cell142],[Cell143],[Cell144],[Cell145],[Cell146],[Cell147],[Cell148],[Cell149],[Cell150]
    )
    SELECT 
        [createdtime],    -- 映射到ReportedTime
        [createdtime],    -- LastChange用源表创建时间填充（也可改为sysdatetime()使用默认值）
        [type],           -- 映射到Type
        [PH],             -- 映射到PH
        -- 以下是源表cell1-cell150字段（与目标表一一对应）
        [cell1],[cell2],[cell3],[cell4],[cell5],[cell6],[cell7],[cell8],[cell9],[cell10],
        [cell11],[cell12],[cell13],[cell14],[cell15],[cell16],[cell17],[cell18],[cell19],[cell20],
        [cell21],[cell22],[cell23],[cell24],[cell25],[cell26],[cell27],[cell28],[cell29],[cell30],
        [cell31],[cell32],[cell33],[cell34],[cell35],[cell36],[cell37],[cell38],[cell39],[cell40],
        [cell41],[cell42],[cell43],[cell44],[cell45],[cell46],[cell47],[cell48],[cell49],[cell50],
        [cell51],[cell52],[cell53],[cell54],[cell55],[cell56],[cell57],[cell58],[cell59],[cell60],
        [cell61],[cell62],[cell63],[cell64],[cell65],[cell66],[cell67],[cell68],[cell69],[cell70],
        [cell71],[cell72],[cell73],[cell74],[cell75],[cell76],[cell77],[cell78],[cell79],[cell80],
        [cell81],[cell82],[cell83],[cell84],[cell85],[cell86],[cell87],[cell88],[cell89],[cell90],
        [cell91],[cell92],[cell93],[cell94],[cell95],[cell96],[cell97],[cell98],[cell99],[cell100],
        [cell101],[cell102],[cell103],[cell104],[cell105],[cell106],[cell107],[cell108],[cell109],[cell110],
        [cell111],[cell112],[cell113],[cell114],[cell115],[cell116],[cell117],[cell118],[cell119],[cell120],
        [cell121],[cell122],[cell123],[cell124],[cell125],[cell126],[cell127],[cell128],[cell129],[cell130],
        [cell131],[cell132],[cell133],[cell134],[cell135],[cell136],[cell137],[cell138],[cell139],[cell140],
        [cell141],[cell142],[cell143],[cell144],[cell145],[cell146],[cell147],[cell148],[cell149],[cell150]
    FROM [RecordSystem].[dbo].[SourceData];

    -- 提交事务（所有数据插入成功后执行）
    COMMIT TRANSACTION;
    PRINT '数据迁移成功！共插入 ' + CAST(@@ROWCOUNT AS VARCHAR) + ' 条记录';
END TRY
BEGIN CATCH
    -- 回滚事务（插入失败时执行）
    ROLLBACK TRANSACTION;
    PRINT '数据迁移失败！错误信息：';
    PRINT '错误号：' + CAST(ERROR_NUMBER() AS VARCHAR);
    PRINT '错误描述：' + ERROR_MESSAGE();
END CATCH;