package com.thopham.projects.dknec.demo_v2

import org.springframework.cloud.openfeign.FeignClient
import org.springframework.stereotype.Component
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestBody

data class AddTagsParams(
        val tags: List<Tag>
)
@Component
@FeignClient(value = "mes.5web.vn", url = "http://mes.5web.vn")
interface MesClient {
    @GetMapping("/manufacturing/getLastTag")
    fun getLastTag(): Tag?
    @PostMapping("/manufacturing/addTags")
    fun addTags(@RequestBody params: AddTagsParams): List<Int>
    @GetMapping("/manufacturing/getFirstTags")
    fun getFirstTags(): List<Tag>
}