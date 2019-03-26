use taikosongsinfo

create table songsinfo(
	songid nchar(20) primary key not null,
	songname nchar(50) not null,
	liang decimal(4) default 0,
	ke decimal(4) default 0,
	buke decimal(4) default 0,
	score decimal(7) default 0,
	lianda decimal(4) default 0,
)

select * from songsinfo

insert into songsinfo values('dsprac','数据库练习曲',1234,567,89,1012340,56)
update songsinfo set songid='yuugen',songname='幽玄ノ乱' where liang=1024

